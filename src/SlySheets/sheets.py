import re
from datetime import datetime, timedelta, tzinfo, timezone
from dataclasses import dataclass
from typing import Any, Callable, Generic, TypeVar
# from enum import Enum

import pytz

from SlyAPI import *

RE_A1 = re.compile(
    # like: 'page'!A1:B2
    r"(?:'(?P<page>\w[\w\s]+\w)'!)?" # TODO: is \w exactly what sheets allows?
    # can be just start_col (e.g. 'A'), or just cell (e.g. 'A1')
    r"(?P<start_col>[a-z]{1,2})(?P<start_row>\d+)?" 
    # ranges (e.g. 'A1:B2') or column ranges (e.g. 'A:B')
    r"(?:\:(?P<end_col>[a-z]{1,2})(?P<end_row>\d+)?)?", 
    re.IGNORECASE
)
RE_LETTERS = re.compile(r'[a-z]+', re.IGNORECASE)

ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
LOTUS123_EPOCH = datetime(1899, 12, 30, 0, 0, 0, 0, timezone.utc)

CellValue = int|str|float

class Scope: # (Enum):
    SheetsReadOnly = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    Sheets         = 'https://www.googleapis.com/auth/spreadsheets'

class ValueRenderOption(EnumParam):
    '''
    What format to return cells in, since formula and display content do not always match.
    From https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption.
    '''
    FORMATTED = 'FORMATTED_VALUE'
    PLAIN     = 'UNFORMATTED_VALUE'
    FORMULA   = 'FORMULA'

class ValueInputOption(EnumParam):
    '''
    Whether to interpret the new value literally, or to parse the same as a user typing it in.
    From https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption.
    '''
    RAW       = 'RAW'
    USER      = 'USER_ENTERED'

class InsertDataOption(EnumParam):
    '''
    Whether to insert new rows when updating a range, or overwrite any existing values.
    From https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append#InsertDataOption.
    '''
    OVERWRITE = 'OVERWRITE'
    INSERT    = 'INSERT_ROWS'

class MajorDimension(EnumParam):
    '''
    Whether elements of the returned value array are columns or rows.
    From https://developers.google.com/sheets/api/reference/rest/v4/Dimension.
    '''
    ROW       = 'ROWS'
    COLUMN    = 'COLUMNS'

class DateTimeRenderOption(EnumParam):
    '''
    How to format cells containing date/time/datetimes.
    From https://developers.google.com/sheets/api/reference/rest/v4/DateTimeRenderOption.
    '''
    Lotus123      = 'SERIAL_NUMBER'
    Formatted     = 'FORMATTED_STRING'

def indexToCol(index: int) -> str:
    '''
    Convert a number to it's corresponding A1 notation column letter(s).
    '''
    a1 = ''
    while index > -1:
        a1 += ALPHABET[index%26]
        index -= 26
    return a1

def colToIndex(letters: str) -> int:
    '''
    Convert A1 notation column letter(s) to it's corresponding index number.
    '''
    letters = letters.upper()
    if len(letters) == 0:
        raise ValueError('Empty column name')
    elif len(letters) == 1:
        return ALPHABET.index(letters)
    else:
        index = 0
        for power, letter in enumerate(reversed(letters)):
            index += ALPHABET.index(letter)*pow(26, power)
        return index

T = TypeVar('T')
U = TypeVar('U')

def maybe(f: Callable[[T], U], x: T) -> U|None:
    '''
    Return f(x) if x is not None, else None.
    '''
    if x is not None:
        return f(x)
    else:
        return None

@dataclass
class GridProperties:
    '''
    Properties of a sheet.
    '''
    frozenColumnCount: int
    frozenRowCount: int
    columnCount: int

@dataclass
class SheetMetadata:
    '''
    Metadata about a sheet.
    '''
    title: str
    timezone: tzinfo


@dataclass
class CellRange:
    '''
    0-indexed, inclusive, integer bounds for a range
    None means unbounded
    '''
    page: str
    from_row: int
    to_row: int|None
    from_col: int
    to_col: int

    def __str__(self):
        '''Convert to A1 Notation'''
        from_colA1 = indexToCol(self.from_col)
        to_colA1 = indexToCol(self.to_col)
        s = F"'{self.page}'!{from_colA1}"
        
        # single cell
        if self.to_col == self.from_col and self.to_row == self.from_row:
            return s + F'{self.from_row+1}'
        # whole columns
        elif self.from_row == 0 and self.to_row is None:
            return s + F':{to_colA1}'
        # bottoms of columns
        elif self.to_row is None:
            return s + F'{self.from_row+1}:{to_colA1}'
        else:
            return s + F'{self.from_row+1}:{to_colA1}{self.to_row+1}'

    def __repr__(self):
            return F"<CellRange at {self}>"

    def columns(self):
        return indexToCol(self.from_col), indexToCol(self.to_col)


def isLetters(t: str):
    return RE_LETTERS.fullmatch(t) is not None

def sheets_date(timestamp: float | int, tz: tzinfo) -> datetime:
    '''
    Convert a DateTimeRenderOption.Lotus123 formatted datetime value to a python datetime object.
    '''
    return LOTUS123_EPOCH.astimezone(tz)+timedelta(days=timestamp)

# T = TypeVar('T')

class ResultRow(Generic[T]):
    '''
    Wraps a list. Allows for indexing by the name of the header.
    '''
    def __init__(self, values: list[T], spreadsheet: 'Sheet'):
        self.values = values
        if spreadsheet.page is not None:
            self.headers = spreadsheet.headers.get(spreadsheet.page, [])
        else:
            self.headers = []

    def __getitem__(self, key: slice|str|int):
        match key:
            case str():
                if key in self.headers:
                    return self.values[self.headers.index(key)]
                elif isLetters(key):
                    return self.values[colToIndex(key)]
                else:
                    raise KeyError(F"Column header was not specified or does not exist: {key}")
            case _:
                return self.values[key]

    def __iter__(self):
        return iter(self.values)

    def __str__(self):
        return str(self.values)

class Sheet(WebAPI):
    """
    Class for handling sheets
    """
    base_url = "https://sheets.googleapis.com/v4/spreadsheets"
    id: str
    page: str|None # TODO: infer?
    n_columns: int
    headers: dict[str, list[str]] # { page: [header1, header2, ...], ... }
    title: str
    timezone: tzinfo

    def __init__(self, app: OAuth2 | dict[str, str], user: OAuth2User | str, sheet_id: str, page_name: str|None=None):
        if isinstance(user, str):
            user = OAuth2User(user)

        if isinstance(app, dict):
            auth = OAuth2(app, user)
        else:
            auth = app
            auth.user = user
        super().__init__(auth)
        self.id = sheet_id
        self.page = page_name

    async def _async_init(self):
        '''
        Fetch the sheet's metadata and optionally headers.
        '''
        await super()._async_init()
        metadata = await self._spreadsheets_get()
        self.title = metadata['properties']['title']
        self.tz = pytz.timezone(metadata['properties']['timeZone'])

        if self.page is not None:
            sheets: list[dict[str, Any]] = metadata['sheets']
            default_sheet_data = None
            for sheet in sheets:
                if sheet['properties']['title'] == self.page:
                    default_sheet_data = sheet
                    break
            if default_sheet_data is None:
                raise ValueError(F"Sheet {self.page} does not exist")
            self.n_columns = default_sheet_data['properties']['gridProperties']['columnCount']
        self.headers = {}
        if self.page is not None:
            header_row = await self[0]
            assert isinstance(header_row, ResultRow)
            self.headers = {
                self.page: list(header_row)
            }


    async def parse_cell_range(self,
            key: int|str|slice|tuple[int|str|slice, int|str|slice],
            default_col_range: tuple[int, int],
            ) -> CellRange:
        '''
        Turn any access specifier into an abstract representation.
        Turning A1 notation into this and back allows for some validation.
        '''

        # defaults
        start_col, end_col = default_col_range
        start_row = 0
        end_row = None
        if self.page is not None:
            page = self.page
        else:
            page = None

        match key:
            case int(): # single row
                start_row = end_row = key
            case slice(): # row range
                if key.step is not None and key.step != 1:
                    raise ValueError(F"Step size must be 1 for row ranges: {key}")
                start_row = key.start

                # CellRange in inclusive, python slices are not
                end_row = key.stop-1 if key.stop is not None else None

            case str() if m := RE_A1.fullmatch(key): # A1 notation
                page = m.group('page') or page
                start_row = maybe(lambda x: int(x)-1, m.group('start_row')) or start_row
                end_row = maybe(lambda x: int(x)-1, m.group('end_row')) or end_row
                start_col = maybe(colToIndex, m.group('start_col')) or start_col
                end_col = maybe(colToIndex, m.group('end_col')) or end_col
            case str(): # header / single column
                if page is None:
                    raise ValueError(F"Cannot specify a column by header without a page name: {key}")
                start_col = end_col = self.headers[page].index(key)
            case (x, y): # 2d index
                match x:
                    case int(): # single row
                        start_row = end_row = x
                    case slice(): # row range
                        start_row = x.start
                        end_row = x.stop-1 if x.stop is not None else None
                    case _:
                        raise ValueError(F"Invalid row index: {x}")
                match y:
                    case int(): # single column
                        start_col = end_col = y
                    case slice(): # column range
                        start_col = y.start
                        if y.stop is None:
                            raise NotImplemented("Spilling column ranges are not yet supported")
                        end_col = y.stop-1
                    case str(): # header or A1 column
                        if page is None:
                                raise ValueError(F"Cannot specify a column by header without a page name: {key}")
                        try:
                            start_col = end_col = self.headers[page].index(y)
                        except ValueError:
                            if RE_LETTERS.fullmatch(y):
                                start_col = end_col = colToIndex(y)
                            else:
                                raise ValueError(F"Invalid column index: {y}")
                    case _:
                        raise ValueError(F"Invalid column index: {y}")
            case _:
                raise ValueError(F"Invalid index: {key}")
        if page is None:
            raise ValueError(F"Could not determine page for range: {key}")
        return CellRange(
                page, start_row, end_row, start_col, end_col)

    async def set(self, range_: str, values: list[list[CellValue]]|list[CellValue]|CellValue):

        range = await self.parse_cell_range(range_, (0, self.n_columns-1))

        # promote low-dimension lists to 2d
        match values:
            case int()|str()|float():
                values = [[values]]
            case list() if not isinstance(values[0], list):
                values = [values] # type: ignore
            case _: pass
        
        await self._values_update(
            str(range),
            {
                'range': str(range),
                'majorDimension': 'ROWS',
                'values': values
            })

    async def date_at(self, range_: str) -> datetime:
        stamp = await self[range_]
        if not isinstance(stamp, (float, int)):
            raise ValueError(F"Expected a sheets timestamp, got {stamp}")
        return sheets_date(stamp, self.tz)

    async def __getitem__(self, *keys: int|str|slice|tuple[int|slice, int|str|slice]) -> float | int | str | ResultRow[Any] | list[ResultRow[Any]]:
        key: int|str|slice|tuple[int|slice, int|str|slice]
        if len(keys) == 1:
            key = keys[0]
        elif len(keys) == 2:
            key = (keys[0], keys[1]) # type: ignore
        else:
            raise ValueError(F"Invalid number of keys: {keys}")

        range = await self.parse_cell_range(key, (0, self.n_columns-1))

        values = (await self._values_get(str(range)))['values']

        # reduce dimensions, for convenience
        # TODO: maybe only when not slicing/ A1 range
        if len(values) == 1 and len(values[0]) == 1:
            return values[0][0]
        elif len(values) == 1 and not isinstance(key, slice):
            return ResultRow(values[0], self)
        else:
            return [ResultRow(row, self) for row in values]

    # async def getCells(self, key, cols=None) -> Any:

    #     shortcut, rangef, step = await self.parseCellRange(key, True, cols)

    #     shortstop = rangef.endRow
    #     if shortstop is not None: shortstop += 1
        
    #     if shortcut:
    #         values = (await self.rows(rangef.columns()))[rangef.startRow:shortstop]
    #     else:
    #         values = (await self._values_get(rangef))['values']

    #     if step > 1:
    #         values = values[slice(None, None, step)]

    #     # reduce dimensions, for convenience
    #     if len(values) == 1 and len(values[0]) == 1:
    #         return values[0][0]
    #     elif len(values) == 1 and not isinstance(key, slice):
    #         return ResultRow(values[0], self)
    #     else:
    #         return [ResultRow(row, self) for row in values]
    
    async def append(self, row: CellValue|list[CellValue]|None=None, search_: str = 'A1', **kwargs: CellValue):
        if row is None and not kwargs:
            raise ValueError("Must provide data to append")
        elif row is None:
            if self.page is None:
                raise ValueError("A default page must be specified for appending into rows by header.")
            if not all(k in self.headers[self.page] for k in  kwargs):
                raise ValueError("Unrecognized header keyword(s)")
            elif self.page is None:
                raise ValueError("Cannot append to a sheet without a page name")
            else:
                row = [kwargs[key] for key in self.headers[self.page]]
        if isinstance(row, (int, str, float)):
            row = [row]
        if not '!' in search_ and self.page:
            search_range = self.page+'!'+search_
        else:
            search_range = search_
        await self._values_append(
            search_range,
            {
                'range': search_range,
                'majorDimension': 'ROWS',
                'values': [row]
            })

    # async def getSpreadSheetInfo(self, id: str) -> 'Spreadsheet':
    #     d = await self._spreadsheets_get(id)
    #     title = d['title']
    #     timezone = pytz.timezone(d['timeZone'])
    #     return title, timezone

    async def delete(self, *keys: int|str|slice):
        if len(keys) == 1:
            key = keys[0]
        else:
            key = (keys[0],  keys[1])
        # if isinstance(key, int) and key < 0: # indexing from end
        #     key = self.n_rows + key
        # elif isinstance(key, slice) and key.start < 0:
        #     key = slice(self.n_rows + key.start, key.stop, key.step)
        range = await self.parse_cell_range(key, (0, self.n_columns-1))
        await self._values_clear(str(range))

    async def _values_get(self, range: str) -> dict[str, Any]:
        return await self.get_json(
            F"/{self.id}/values/{range}",
            ValueRenderOption.PLAIN.to_dict()
        )

    async def _values_append(self, range: str, value: dict[str, Any]) -> dict[str, Any]:
        return await self.post_json(
            F"/{self.id}/values/{range}:append",
            params=ValueInputOption.RAW.to_dict(), 
            json=value)

    async def _values_clear(self, range: str) -> dict[str, Any]:
        return await self.post_json(
            F"/{self.id}/values/{range}:clear"
            )

    async def _values_update(self, range: str, value: dict[str, Any]) -> dict[str, Any]:
        return await self.put_json(
            F"/{self.id}/values/{range}",
            params=ValueInputOption.RAW.to_dict(), 
            json=value)

    async def _spreadsheets_get(self, sheet_id: str|None=None) -> dict[str, Any]:
        if sheet_id is None: sheet_id = self.id
        return await self.get_json(
            F"/{sheet_id}",
        )

    # async def rowCount(self) -> int:
    #     return len(await self.rows())

    # async def rows(self, cols=None) -> list[list[Any]]:
    #     n, rows = 0, []
    #     if cols is None: cols = self.cols
    #     while len(result := await self.getCells(slice(n, n+500, 1), cols=cols)) > 0:
    #         n += 500
    #         rows += result
    #         if len(result) < 500: break
    #     return rows




    

