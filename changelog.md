# Changelog

This file contains all the notable changes done to the Ballerina `xlsx` package through the releases.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- File-based XLSX parsing via `parse()` to `record{}[]`, `map<anydata>[]`, or `string[][]`
- File-based XLSX writing via `write()` from records, maps, or 2D arrays
- Excel Table operations via `parseTable()` and `writeTable()` (auto-expands when data exceeds current table size)
- `Workbook` class via factory functions `openFile(path)`, `createFile(path)`, and `createWorkbook()`, with `getSheetNames`, `getSheetCount`, `getSheet`, `getSheetByIndex`, `createSheet`, `deleteSheet`, `deleteSheetByIndex`, `save`, `saveAs`, `close`, `getTable`, and `getAllTables`
- `Sheet` class with `getName`, `getUsedRange`, `getUsedCellRange`, `getRowCount`, `getColumnCount`, `getRows`, `getRow`, `putRows`, `getTable`, `getTables`, `createTable`, `createTableFromData`, and `deleteTable`
- `Table` class with `getName`, `getDisplayName`, `getSheetName`, `getRange`, `getDataRange`, `getRowCount`, `getColumnCount`, `getHeaders`, `getRows`, `getRow`, `putRows`, `hasTotalsRow`, `getTotalsRow`, `rename`, and `resize`
- `@xlsx:Name` annotation for bidirectional header-to-field mapping (used on parse and write)
- `Row` wrapper type for preserving original row positions during round-trip operations (empty rows retain their `rowIndex` with `value = null`)
- `FormulaMode` enum (`CACHED`, `TEXT`) for formula cell handling
- Header-less parsing via `headerRowIndex = ()` — columns become `col0`, `col1`, etc.
- Case-insensitive header matching via `caseInsensitiveHeaders` option
- Row count limit via `rowCount` option for previewing large files
- Data projection via `allowDataProjection` (`nilAsOptionalField`, `absentAsNilableType`) with strict mode (`false`)
- Constraint validation via `enableConstraintValidation` (integrates with Ballerina `@constraint` annotations)
- Fail-safe error handling via `FailSafeOptions` — continue parsing on row-level errors with console and/or file logging (`METADATA`, `RAW`, `RAW_AND_METADATA` content types; `APPEND` or `OVERWRITE` file modes)
- Used-range detection that excludes formatted-but-empty "ghost rows"
- `CellRange` type for representing rectangular sheet regions
- Distinct error types: `Error`, `ParseError`, `FileNotFoundError`, `SheetNotFoundError`, `TypeConversionError`, `ConstraintValidationError`, `TableNotFoundError`, `TableOverlapError`, `InvalidTableRangeError`
- `ErrorDetails` record carrying `sheetName`, `tableName`, `cellAddress`, `rowNumber`, and `columnNumber` context

### Reserved
- `parseAsStream()` API signature reserved for future SAX-based streaming support
