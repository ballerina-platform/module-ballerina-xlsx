# Changelog

This file contains all the notable changes done to the Ballerina `xlsx` package through the releases.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- File-based XLSX parsing via `parseSheet()` to `record{}[]`, `map<anydata>[]`, or `string[][]`
- File-based XLSX writing via `writeSheet()` from records, maps, or 2D arrays
- Tier 1 table conveniences: `parseTable(path, tableName, ...)` and `writeTable(data, path, tableName, ...)` for one-shot table-by-name flows (tables are unique by name across the workbook, so no sheet specifier is needed). `writeTable` auto-expands the target table.
- Target-type-driven date / time / date-time binding to `time:Civil`, `time:Date`, and `time:TimeOfDay` from the standard `ballerina/time` module. ISO 8601 `string` is the fallback when the field type is `string` or `anydata`. Date arithmetic correctness fixes ship alongside: `BigDecimal` for time-of-day math (eliminates the 23:59:59.999 → 00:00:00 rollover), `ZoneOffset.UTC` for naive cells (reproducible across machines), and honour `Workbook.isDate1904()` (1904-origin Mac files parse correctly).
- Large-integer write protection: integer values with `|n| > 2^53` are written as text cells containing the exact digit string. Smaller integers continue to write as numeric cells. The fallback is automatic — common 16+ digit account / order IDs now round-trip losslessly without the user having to declare them as `string`.
- `Workbook` class with consolidated `init(string|byte[]? source = ())`:
  - `new` — empty in-memory workbook
  - `new(path)` — open an existing file (errors if missing)
  - `new(bytes)` — open from a byte array
- `Workbook` methods: `getSheetNames`, `getSheetCount`, `hasSheet`, `getSheet`, `createSheet`, `deleteSheet`, `getTable`, `getAllTables`, `save`, `saveAs`, `toBytes`, `close`
- `Sheet` class with row, column, and cell access: `getName`, `getUsedRange`, `getUsedCellRange`, `getRowCount`, `getColumnCount`, `getRows`, `getRow`, `getColumn`, `getCell`, `putRows`, `setRow`, `setColumn`, `setCell` (× 2 overloads — index and A1 notation), `deleteRow`, `rename`
- `Sheet` table management: `getTable`, `getTables`, `createTable`, `createTableFromData`, `deleteTable`
- `Table` class (Excel Tables / ListObjects): `getName`, `getDisplayName`, `getSheetName`, `getRange`, `getDataRange`, `getRowCount`, `getColumnCount`, `getHeaders`, `getRows`, `getRow`, `putRows` (auto-expands), `hasTotalsRow`, `getTotalsRow`, `rename`, `resize`
- `@xlsx:Name` annotation for bidirectional header-to-field mapping (read and write)
- `FormulaMode` enum (`CACHED`, `TEXT`) for read-only formula cell handling
- Headerless parsing via `headerRowIndex = ()` — columns become `col0`, `col1`, etc.
- Case-insensitive header matching via `caseInsensitiveHeaders` option
- Row count limit via `rowCount` option for previewing large files
- Data projection via `allowDataProjection` (`nilAsOptionalField`, `absentAsNilableType`) with strict mode (`false`)
- Constraint validation via `enableConstraintValidation` (integrates with Ballerina `@constraint` annotations)
- Fail-safe error handling via `FailSafeOptions` — continue parsing on row-level errors with console and/or file logging (`METADATA`, `RAW`, `RAW_AND_METADATA` content types; `APPEND` or `OVERWRITE` file modes)
- Atomic file writes — `writeSheet`, `Workbook.save`, and `Workbook.saveAs` use a sibling temp file and atomic rename so a failed write never destroys the original file
- Phantom-reference leak protection — workbooks that escape without `close()` are reclaimed by a background cleanup thread
- `CellRange` type for representing rectangular sheet regions (0-based indices)
- Distinct error types: `Error`, `ParseError`, `FileNotFoundError`, `SheetNotFoundError`, `TypeConversionError`, `ConstraintValidationError`, `TableNotFoundError`, `TableOverlapError`, `InvalidTableRangeError`
- `ErrorDetails` record carrying `sheetName`, `tableName`, `cellAddress`, `rowNumber` (1-based), and `columnNumber` (1-based) context

### Not in v0.1

- Formula authoring on write — strings starting with `=` are written verbatim as text. A `Formula` wrapper type is deferred to a later release.
- Formula re-evaluation — no `EVALUATE` / `RECALCULATE` / `PRESERVE` modes.
- Streaming — no row-streaming API for files larger than memory; no byte-streaming `parseStream`.
- Top-level `parseBytes` / `writeToBytes` — use `new Workbook(byte[])` and `Workbook.toBytes()`.
- `*xlsx:Row` round-trip spread — every row in the used range is materialised by default; round-trip preservation that needs more than cell values uses the Workbook API.
- `WriteOptions.strictNumericPrecision` flag — v0.1 silently falls back to text cells for large ints. A hard-error flag is a possible later addition.
- XLS (legacy 97-2003) format, password-protected files, named ranges, cell styling, and range operations.
