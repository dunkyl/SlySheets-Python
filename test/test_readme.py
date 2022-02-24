import json
from SlySheets import *

async def test_readme():

    app = json.load(open('test/app.json'))

    sheet = await Sheet(app, 'test/user.json', '1arnulJxyi-I6LEeCPpEy6XE5V87UF54dUAo9F8fM5rw', 'Sheet 1')

    print(sheet.title)
    print(sheet.headers)

    # A1 notation
    a1 = await sheet['A1']
    print(F"Cell A1: {a1}") # Cell A1: Foo

    # or zero-indexed row/col
    first_cell = await sheet[0, 0]
    print(F"Cell A1 (again): {first_cell}") # Cell A1 (again): Foo

    # zero-indexed rows
    first_row = await sheet[0]
    assert isinstance(first_row, ResultRow)
    print(F" | {first_row[0]:6} | {first_row[1]:6} |") # | Foo     | Bar     |

    # header-indexed columns
    foos = await sheet['Foo'] # list[Any]
    print(F"Foos: {foos}") # Foos: [1, 2, 3, 26]

    # slicing and iterators
    # TODO: async iters like ytdapi
    for row in await sheet[1:3]:
        # index result by header
        print(F" | {row['Foo']:6} | {row['Bar']:6} |") # |      1 | a     | etc...
    for row in await sheet[3:]:
        # or by column
        print(F" | {row['A']:6} | {row['B']:6} |") # |    26 | z     | etc...

    # append, of course
    await sheet.append([0, 'x'])
    await sheet.append(Foo=1, Bar='y')

    # TODO: consider not using slices to simplify    
    # await sheet.delete(slice(-2, None))
    await sheet.delete('A6:B7')
    await sheet.set("'Sheet 1'!E3", 'Hello World!')

    # dates
    today = await sheet.date_at('D5') # =TODAY()
    print(F"It is now {today.isoformat()} (timezone: {sheet.tz})")