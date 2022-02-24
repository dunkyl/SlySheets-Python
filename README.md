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

You can directly grant user tokens using the command line, covering the whole OAuth 1.0 grant process.

---

Example usage:

```py
import asyncio
from SlySheets import *

async def main():
    auth = OAuth2User.from_files('client.json', 'user.json')
    my_sheet = await Sheet(' < your sheet id > ', auth)

    print(sheet.title)

    # A1 notation
    a1 = await sheet['A1']
    print(F"Cell A1: {a1}")

    # or zero-indexed row/col
    first_cell = await sheet[0, 0]
    print(F"Cell A1 (again): {first_cell}")

    # zero-indexed rows
    first_row = await sheet[0]
    print(F" | {first_row[0]:8} | {first_row[1]:8} |")

    # header-indexed columns
    foos = await sheet['Foo'] # list[Any]
    print(F"Foos: {foos}")

    # slicing and async iterators
    async for row in sheet[1:3]:
        # index result by header
        print(F" | {row['Foo']:8} | {row['Bar']:8} |")
    async for row in sheet[3:]:
        # or by column
        print(F" | {row['A']:8} | {row['B']:8} |")

    # append, of course
    await sheet.append([0, 'x'])
    await sheet.append(Foo=1, Bar='y')
    
    await sheet.delete[-2:]
    await sheet.set['Sheet 1!E3']('Hello World!')

    # dates
    today = await sheet.date_at('D5') # =TODAY()
    print(F"It is now {today.isoformat()} (timezone: {sheet.tz})")

asyncio.run(main())
```
