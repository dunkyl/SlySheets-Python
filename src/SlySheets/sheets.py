'''
Google Sheets API and types.
https://developers.google.com/sheets/api/guides/concepts
'''
from enum import Enum
import re
from datetime import datetime, timedelta, tzinfo, timezone
from dataclasses import dataclass
from typing import Any, TypeVar

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
# note: timezone here is a workaround on windows for OSERROR 22, since the lotus epoch predates 1970
LOTUS123_EPOCH = datetime(1899, 12, 30, 0, 0, 0, 0, timezone.utc)

CellValue = int|str|float|None

class Scope: # (Enum):
    SheetsReadOnly = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    Sheets         = 'https://www.googleapis.com/auth/spreadsheets'

class ValueRenderOption(Enum):
    '''
    What format to return cells in, since formula and display content do not always match.
    From https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption.
    '''
    FORMATTED = 'FORMATTED_VALUE'
    PLAIN     = 'UNFORMATTED_VALUE'
    FORMULA   = 'FORMULA'

class ValueInputOption(Enum):
    '''
    Whether to interpret the new value literally, or to parse the same as a user typing it in.
    From https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption.
    '''
    RAW       = 'RAW'
    USER      = 'USER_ENTERED'

class InsertDataOption(Enum):
    '''
    Whether to insert new rows when updating a range, or overwrite any existing values.
    From https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append#InsertDataOption.
    '''
    OVERWRITE = 'OVERWRITE'
    INSERT    = 'INSERT_ROWS'

class MajorDimension(Enum):
    '''
    Whether elements of the returned value array are columns or rows.
    From https://developers.google.com/sheets/api/reference/rest/v4/Dimension.
    '''
    ROW       = 'ROWS'
    COLUMN    = 'COLUMNS'

class DateTimeRenderOption(Enum):
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

class CellRange:
    '''
    0-indexed, inclusive, integer bounds for a range
    None means unbounded
    '''
    page: str | None
    from_row: int
    from_col: int
    to_row: int | None # None is case of unbounded
    to_col: int

    def __init__(self, a1: str):
        '''
        Initialise from A1 notation.
        '''
        match = RE_A1.match(a1)
        if match is None:
            raise ValueError(F"Invalid A1 notation: {a1}")
        self.page = match['page']
        self.from_row = int(match['start_row']) - 1
        if match['end_col']:
            if match['end_row']:
                self.to_row = int(match['end_row']) - 1
            else:
                self.to_row = None
        else:
            self.to_row = self.from_row
        self.from_col = colToIndex(match['start_col'])
        self.to_col = colToIndex(match['end_col']) if match['end_col'] else self.from_col

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

    def shape(self):
        '''
        Return the number of rows and columns in the range.
        '''
        if self.to_row is None:
            return (self.to_col - self.from_col + 1, 0)
        else:
            return (self.to_col - self.from_col + 1, self.to_row - self.from_row + 1)

def sheets_date(timestamp: float | int, tz: tzinfo) -> datetime:
    '''
    Convert a DateTimeRenderOption.Lotus123 formatted datetime value to a python datetime object.
    '''
    return LOTUS123_EPOCH.astimezone(tz)+timedelta(days=timestamp)

class Page:
    _sheet: 'Spreadsheet'
    id: int
    title: str
    n_columns: int

    def __init__(self, page_meta: dict[str, Any], sheet: 'Spreadsheet'):
        page_props = page_meta['properties']
        self._sheet = sheet
        self.id = page_props['sheetId']
        self.title = page_props['title']
        self.n_columns = page_props['gridProperties']['columnCount']

    def link(self):
        return F'{self._sheet.link()}#gid={self.id}'

    async def range(self, a1: str):
        a1_ = CellRange(a1)
        if a1_.page is None:
            a1_.page = self.title
        return await self._sheet.range(a1_)

    async def delete_range(self, a1: str):
        a1_ = CellRange(a1)
        if a1_.page is None:
            a1_.page = self.title
        return await self._sheet.delete_range(a1_)

    async def set_range(self, a1: str, values: list[list[Any]]):
        a1_ = CellRange(a1)
        if a1_.page is None:
            a1_.page = self.title
        return await self._sheet.set_range(a1_, values)

    async def set_cell(self, a1: str, value: list[list[Any]]):
        return await self.set_range(a1, [[value]])

    async def rows(self, start: int, end: int, n_cols: int | None = None):
        '''Get the content of a range of rows, inclusive, zero-indexed'''
        n_cols = n_cols or self.n_columns
        a1 = F'A{start+1}:{indexToCol(n_cols-1)}{end+1}'
        return (await self.range(a1))

    async def row(self, index: int, n_cols: int | None = None):
        '''Get the content of a single row, zero-indexed'''
        return (await self.rows(index, index, n_cols))[0]

    async def cell(self, a1: str):
        '''Get the content of a single cell'''
        return (await self.range(a1))[0][0]

    async def date_at_cell(self, a1: str) -> datetime:
        '''Get the date at a cell'''
        a1_ = CellRange(a1)
        if a1_.page is None:
            a1_.page = self.title
        return await self._sheet.date_at_cell(str(a1_))

    async def column(self, index: int):
        a1_col = indexToCol(index)
        return [ r[0] for r in await self.range(F'{a1_col}1:{a1_col}')]

    async def column_named(self, name: str):
        header_row = await self.row(0)
        if name in header_row:
            return await self.column(header_row.index(name))
        else:
            raise KeyError(F"Column header was not specified or does not exist: {name}")

    async def rows_dicts(self, start: int, end: int):
        header_row = [str(h) for h in await self.row(0)]
        rows = [ dict(zip(header_row, row)) for row in await self.rows(start, end) ]
        return rows

    async def extend(self, values: list[list[CellValue]], search_: str = 'A1'):
        await self._sheet.extend(values, self.title, search_)

    async def extend_dicts(self, values: list[dict[str, CellValue]], search_: str = 'A1'):
        header_row = [str(h) for h in await self.row(0)]
        rows = [ [obj.get(h, None) for h in header_row] for obj in values ]
        await self._sheet.extend(rows, self.title, search_)

    async def append(self, row: list[CellValue], search_: str = 'A1'):
        await self.extend([row], search_)

    async def append_dict(self, obj: dict[str, CellValue], search_: str = 'A1'):
        await self.extend_dicts([obj], search_)

    def batch(self):
        return BatchEdit(self)

@dataclass
class BatchEditOp:
    kind: str
    content: dict[str, Any]

class BatchEdit:
    _page: Page
    _requests: list[BatchEditOp]

    # TODO: finish
    def __init__(self, page: Page):
        self._page = page
        self._requests = []

    async def __aenter__(self):
        return self
    
    async def __aexit__(self, _exc_type: Any, _exc_val: Any, _exc_tb: Any):

        # TODO: call
        # POST https://sheets.googleapis.com/v4/spreadsheets/{spreadsheetId}:batchUpdate
        # {
        #     "requests": [
        #         {
        #         object (Request)
        #         }
        #     ],
        #     "includeSpreadsheetInResponse": boolean,
        #     "responseRanges": [
        #         string
        #     ],
        #     "responseIncludeGridData": boolean
        #     }
        pass

    def set_range(self, a1: str, values: list[list[CellValue]]):
        field_mask = "userEnteredValue"
        self._requests.append(BatchEditOp(
            'UpdateCellsRequest', {
                'rows': [
                    { 'values': [ { 'userEnteredValue': v } for v in row ] } for row in values
                ],
                'fields': field_mask,
                'range': str(a1)
            }))

class Spreadsheet(WebAPI):
    """
    Class for handling sheets
    """
    base_url = "https://sheets.googleapis.com/v4/spreadsheets"
    id: str

    _tz = None

    def __init__(self, auth: OAuth2, sheet_id: str):
        super().__init__(auth)
        self.id = sheet_id

    def link(self):
        '''Hyperlink to the spreadsheet'''
        return F'https://docs.google.com/spreadsheets/d/{self.id}/edit#gid=0'

    async def title(self):
        '''Title of the spreadsheet'''
        return (await self._spreadsheets_get())['properties']['title']
    
    # get last tz if already fetched
    async def _timezone(self):
        if self._tz is None:
            return await self.tz()
        return self._tz

    async def tz(self):
        '''Default timezone of the spreadsheet'''
        tz_str = (await self._spreadsheets_get())['properties']['timeZone']
        self._tz = pytz.timezone(tz_str)
        return self._tz
    
    async def pages(self) -> list[Page]:
        '''Get a `Page` for each page in the spreadhsheet'''
        return [
            Page(page_meta, self)
            for page_meta
            in (await self._spreadsheets_get())['sheets']
        ]

    async def page(self, title: str):
        '''Get a `Page` by title'''
        pages = await self.pages()
        for page in pages:
            if page.title == title:
                return page
        raise KeyError(F"Page {title} does not exist")

    async def range(self, a1: str | CellRange) -> list[list[CellValue]]:
        if isinstance(a1, str):
            a1 = CellRange(a1)
        if a1.page is None:
            raise ValueError(F"Page not specified: {a1}")

        values = (await self._values_get(str(a1)))['values']

        # fill in omitted cells on bottom and right with None
        shape_x, shape_y = a1.shape()

        print(F"{a1} shape: {shape_x}, {shape_y}")

        if shape_y > len(values):
            values += [[]] * (shape_y - len(values))
        for i, row in enumerate(values):
            if shape_x > len(row):
                values[i] += [None] * (shape_x - len(row))

        return values

    async def delete_range(self, a1: str | CellRange):
        if isinstance(a1, str):
            a1 = CellRange(a1)
        if a1.page is None:
            raise ValueError(F"Page not specified: {a1}")

        await self._values_clear(str(a1))

    async def cell(self, a1: str) -> CellValue:
        return (await self.range(a1))[0][0]

    async def set_range(self, a1: str | CellRange, values: list[list[CellValue]]):
        if isinstance(a1, str):
            a1 = CellRange(a1)
        if a1.page is None:
            raise ValueError(F"Page not specified: {a1}")

        await self._values_update( str(a1), {
                'range': str(a1),
                'majorDimension': 'ROWS',
                'values': values
            })

    async def set_cell(self, a1: str | CellRange, value: CellValue):
        await self.set_range(a1, [[value]])

    async def date_at_cell(self, a1: str) -> datetime:
        stamp = await self.cell(a1)
        if not isinstance(stamp, (float, int)):
            raise ValueError(F"Expected a sheets timestamp, got {stamp}")
        return sheets_date(stamp, await self._timezone())
    
    async def extend(self, values: list[list[CellValue]], page: str, search_: str = 'A1'):
        search_range = F"'{page}'!{search_}"
        await self._values_append(
            search_range,
            {
                'range': search_range,
                'majorDimension': 'ROWS',
                'values': values
            })

    async def _values_get(self, range: str) -> dict[str, Any]:
        return await self.get_json(
            F"/{self.id}/values/{range}",
            { "valueRenderOption": ValueRenderOption.PLAIN }
        )

    async def _values_append(self, range: str, value: dict[str, Any]) -> dict[str, Any]:
        return await self.post_json(
            F"/{self.id}/values/{range}:append",
            { "valueInputOption": ValueInputOption.RAW },
            json=value)

    async def _values_clear(self, range: str) -> dict[str, Any]:
        return await self.post_json(
            F"/{self.id}/values/{range}:clear"
            )

    async def _values_update(self, range: str, value: dict[str, Any]) -> dict[str, Any]:
        return await self.put_json(
            F"/{self.id}/values/{range}",
            { "valueInputOption": ValueInputOption.RAW }, 
            json=value)

    async def _spreadsheets_get(self) -> dict[str, Any]:
        return await self.get_json(F"/{self.id}")

    

