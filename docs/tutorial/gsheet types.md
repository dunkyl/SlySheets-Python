# Google Sheet types

Like with many programming languages, Google Sheets has its own perspective on primitive types


## Type Mapping

| Sheets Type     | Python Type               |
| --------------- | ------------------------- |
| Text            | str                       |
| Number          | int \| float              |
| Date            | datetime                  |
| Boolean         | bool                      |
| Validated       | str                       |
| Range (Value)   | list of list of the above |
| Range (Address) | SlySheets.Range           |