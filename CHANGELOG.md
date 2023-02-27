# Change Log

## [Unreleased]

### Changed
- `Spreadsheet` is no longer `AsyncInit` and should not be awaited
- `Spreadsheet` constructor now takes a `SlyAPI.OAuth2`
- Some methods that were attribtutes have been changed to coroutine functions:
    - `Spreadsheet.pages()`
    - `Spreadsheet.page()`
    - `Spreadsheet.title()`
    - `Spreadsheet.tz()`
- `Spreadsheet.pages()` and `Spreadsheet.page()` are now coroutine functions

---

## [0.1.0] - 2022-02-24
Initial release