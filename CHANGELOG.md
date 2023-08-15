# Change Log

## [Unreleased]

---

## [0.2.3] - 2023-08-14

Hotfix for empty ranges

## [0.2.2] - 2023-08-14

### Added
- `Page.grid_row_count()`

## [0.2.1] - 2023-08-14

### Fixed
- `Spreadsheet.tz()` endless loop

### Added
- `Page.tz()`: same as `Spreadsheet.tz()`

## [0.2.0] - 2023-02-27

### Added
- `Spreadsheet.link()`
- `Page.link()`

### Changed
- `Spreadsheet` is no longer `AsyncInit` and should not be awaited
- `Spreadsheet` constructor now takes a `SlyAPI.OAuth2`
- Some methods that were attributes have been changed to coroutine functions:
    - `Spreadsheet.pages()`
    - `Spreadsheet.page()`
    - `Spreadsheet.title()`
    - `Spreadsheet.tz()`
- `Spreadsheet.pages()` and `Spreadsheet.page()` are now coroutine functions


## [0.1.0] - 2022-02-24
Initial release