# Changelog

This file contains all the notable changes done to the Ballerina `xlsx` package through the releases.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- File-based XLSX parsing via `parseSheet()` to `record{}[]`, `map<anydata>[]`, or `string[][]`
- File-based XLSX writing via `writeSheet()` from records, maps, or 2D arrays
- Tier 1 table conveniences: `parseTable(path, tableName, ...)` and `writeTable(data, path, tableName, ...)` for one-shot table-by-name flows (tables are unique by name across the workbook, so no sheet specifier is needed). `writeTable` auto-expands the target table.
- `Data = Row[]` and `Row = record{} | map<anydata> | string[]` as the public-API row shape. Inline literals (`writeSheet([["A","B"],["1","2"]], path)`) and explicit `Data`/`Row` typing both work end-to-end on read and write paths.
- Target-type-driven date / time / date-time binding to `time:Civil`, `time:Date`, and `time:TimeOfDay` from the standard `ballerina/time` module. ISO 8601 `string` is the fallback when the field type is `string` or `anydata`.
- Cross-machine deterministic datetime round-trips. Serial ↔ LocalDate math is done directly in UTC, bypassing POI's `java.util.Date` conversion (which inherits the system timezone via `LocaleUtil.getLocaleCalendar`). Honours `Workbook.isDate1904()` so legacy Excel-for-Mac files parse correctly. Handles the 1900-epoch Lotus 1-2-3 phantom Feb 29 quirk so date serials match Excel's display.
- Sub-second precision preserved across both read and write paths via `BigDecimal` nano arithmetic. `time:TimeOfDay` round-trips at nanosecond precision; `time:Civil` at microsecond precision (the practical limit of double-precision Excel serials with a date integer component).
- `Workbook` class with consolidated `init(string|byte[]? input = ())`:
  - `new` — empty in-memory workbook
  - `new(path)` — open an existing file (errors if missing)
  - `new(bytes)` — open from a byte array
- `Workbook` methods: `getSheetNames`, `getSheetCount`, `hasSheet`, `getSheet(string|int)`, `createSheet`, `deleteSheet(string|int)`, `getTable`, `getAllTables`, `save`, `saveAs`, `toBytes`, `close`
- `Sheet` as a public `object` type (interface) backed by an internal `SheetImpl` class. Compile-time prevention of direct `new Sheet()` — instances are vended only through `Workbook` methods. Same idiom for `Table`.
- `Sheet` methods: `getName`, `getUsedRange`, `getUsedCellRange`, `getRowCount`, `getColumnCount`, `getRows`, `getRow`, `getColumn`, `getCell`, `putRows`, `setRow`, `setColumn`, `setCell(int, int, anydata)`, `setCellByAddress(string, anydata)` (A1 notation), `deleteRow`, `rename`
- `Sheet` table management: `getTable`, `getTables`, `createTable`, `createTableFromData`, `deleteTable`
- `Table` (Excel Tables / ListObjects): `getName`, `getDisplayName`, `getSheetName`, `getRange`, `getDataRange`, `getRowCount`, `getColumnCount`, `getHeaders`, `getRows`, `getRow`, `putRows` (auto-expands), `hasTotalsRow`, `getTotalsRow`, `rename`, `resize`
- `@xlsx:Name` annotation for bidirectional header-to-field mapping (read and write). Annotation values are trimmed on lookup so accidental whitespace doesn't silently break matching.
- Excel name validation on create and rename:
  - Sheet: 1-31 chars, no `\ / ? * [ ] :`. Sheet name lookups (`getSheet`, `hasSheet`, `deleteSheet`) are case-insensitive, matching Excel's own semantics.
  - Table: 1-255 chars, must start with a letter or underscore, no spaces.
- `Workbook.deleteSheet` refuses to delete the last sheet (Excel rejects opening a sheet-less workbook).
- Duplicate detection at parse time: two sheet columns with the same header → clear error; two record fields with the same `@xlsx:Name` value → clear error.
- Handle invalidation on `Workbook.close()` / `deleteSheet()` / `Sheet.deleteTable()`: any subsequent operation on a vended `Sheet` or `Table` handle returns a typed `Error` rather than panicking on stale native state.
- `FormulaMode` enum (`CACHED`, `TEXT`) for read-only formula cell handling
- Headerless parsing via `headerRowIndex = ()` — columns become `col0`, `col1`, etc.
- Case-insensitive header matching via `caseInsensitiveHeaders` option
- Row count limit via `rowCount` option for previewing large files
- Data projection via `allowDataProjection` (`nilAsOptionalField`, `absentAsNilableType`) with strict mode (`false`)
- Constraint validation via `enableConstraintValidation` (integrates with Ballerina `@constraint` annotations)
- Fail-safe error handling via `FailSafeOptions` — continue parsing on row-level errors with console and/or file logging (`METADATA`, `RAW`, `RAW_AND_METADATA` content types; `APPEND` or `OVERWRITE` file modes)
- Atomic file writes — `writeSheet`, `writeTable`, `Workbook.save`, and `Workbook.saveAs` use a sibling temp file and atomic rename (`Files.move(ATOMIC_MOVE, REPLACE_EXISTING)`) so a failed write never destroys the original file.
- Per-call cell-style cache so concurrent writes don't serialise on a global lock and no workbook-pinning leak in the process-level map.
- Phantom-reference leak protection — workbooks that escape without `close()` are reclaimed by a background cleanup thread (O(1) unregister keyed by `Workbook` identity).
- `CellRange` type for representing rectangular sheet regions (0-based indices)
- Distinct error types: `Error`, `ParseError`, `FileNotFoundError`, `SheetNotFoundError`, `TypeConversionError`, `ConstraintValidationError`, `TableNotFoundError`, `TableOverlapError`, `InvalidTableRangeError`
- `ErrorDetails` record carrying `sheetName`, `tableName`, `cellAddress`, `rowNumber` (1-based), and `columnNumber` (1-based) context

