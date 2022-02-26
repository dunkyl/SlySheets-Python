# ![sly logo](https://raw.githubusercontent.com/dunkyl/SlyMeta/main/sly%20logo.svg) SlySheets for Python

> 🚧 **This library is an early work in progress! Breaking changes may be frequent.**

> 🐍 For Python 3.10+

## No-boilerplate, _async_ and _typed_ Google Sheets access. 😋

```shell
pip install slysheets
```

This library does not have full coverage.
Currently, the following topics are supported:

* Editing sheet cells
* Reading sheet metadata

You can use [SlyAPI](https://github.com/dunkyl/SlyPyAPI) to directly grant user tokens using the command line, covering the whole OAuth 2 user grant process after getting client credentials from [Google](https://console.cloud.google.com/).

---

Example usage:

```py
import asyncio
from SlySheets import *

async def main():

    spreadsheet = await Spreadsheet('test/app.json', 'test/user.json', '1arnulJxyi-I6LEeCPpEy6XE5V87UF54dUAo9F8fM5rw')
    page = spreadsheet.page('Sheet 1')

    # A1 notation
    a1 = await page.cell('A1')
    print(F"Cell A1: {a1}") # Cell A1: Foo

    # zero-indexed rows
    first_row = await page.row(0)
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
    await page.append([0, 'u'])
    await page.append_dict({'Foo': 1, 'Bar': 'v'})

    await page.extend([[3, 'w'], [4, 'x']])
    await page.extend_dicts([{'Foo': 5, 'Bar': 'y'}, {'Foo': 6, 'Bar': 'z'}])

    # TODO: consider not using slices to simplify    
    # await sheet.delete(slice(-2, None))
    await page.delete_range('A6:B11')

    # use spreadsheet object to update a specific page with A1 notation
    await spreadsheet.set_cell("'Sheet 1'!E3", 'Hello World!')

    # dates
    today = await page.date_at_cell('D5') # =TODAY()
    assert isinstance(today, datetime)
    print(F"It is now {today.isoformat()} (timezone: {spreadsheet.tz})")


    # batch edits
    async with sheet.batch() as batch:
        batch.set_range("C2:D3", [[0, 1], [2, 3]])

asyncio.run(main())
```
