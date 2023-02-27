from SlySheets import *

async def test_readme():

    auth = OAuth2('test/app.json', 'test/user.json')

    spreadsheet = Spreadsheet(auth, '1arnulJxyi-I6LEeCPpEy6XE5V87UF54dUAo9F8fM5rw')
    page = await spreadsheet.page('Sheet 1')

    print(page.link()) # https://docs.google.com/spreadsheets/d/1arnulJxyi-I6LEeCPpEy6XE5V87UF54dUAo9F8fM5rw/edit#gid=0

    # A1 notation
    a1 = await page.cell('A1')
    print(F"Cell A1: {a1}") # Cell A1: Foo

    # zero-indexed rows
    first_row = await page.row(0)
    print(first_row)
    print(F" | {first_row[0]:6} | {first_row[1]:6} |") # | Foo     | Bar     |

    # header-indexed columns
    foos = await page.column_named('Foo')
    print(F"Foos: {foos}") # Foos: [1, 2, 3, 26]

    # zero-indexed columns
    foos_2 = await page.column(0)
    assert foos == foos_2

    for row in await page.rows_dicts(1, 4):
        # index result by header
        print(F" | {row['Foo']:6} | {row['Bar']:6} |") # |      1 | a     | etc...

    # append and extend, of course
    await page.append([21, 'u'])
    await page.append_dict({'Foo': 22, 'Bar': 'v'})

    await page.extend([[23, 'w'], [24, 'x']])
    await page.extend_dicts([{'Foo': 25, 'Bar': 'y'}, {'Foo': 26, 'Bar': 'z'}])

    # TODO: consider not using slices to simplify    
    # await sheet.delete(slice(-2, None))
    await page.delete_range('A6:B11')

    # use spreadsheet object to update a specific page with A1 notation
    await spreadsheet.set_cell("'Sheet 1'!E3", 'Hello World!')

    # dates
    today = await page.date_at_cell('D5') # =TODAY()
    assert isinstance(today, datetime)
    print(F"It is now {today.isoformat()} (timezone: {await spreadsheet.tz()})")

    # batch edits
    async with page.batch() as batch:
        batch.set_range("C2:D3", [[0, 1], [2, 3]])