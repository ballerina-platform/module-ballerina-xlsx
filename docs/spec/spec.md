_Owners_: @YasanPunch \
_Reviewers_: @niveathika \
_Created_: 2026/05/02 \
_Updated_: 2026/05/02 \
_Edition_: Swan Lake

# Specification: Ballerina XLSX Module

## Introduction

This is the specification for the `xlsx` module of the [Ballerina language](https://ballerina.io/), which provides functionality for reading and writing Microsoft Excel files in the XLSX (Office Open XML) format with type-safe data binding to Ballerina records.

The `xlsx` module specification is written to describe the functionality available from version 1.0.0 onwards.

If you have any feedback or suggestions about the module, start a discussion via a [GitHub issue](https://github.com/ballerina-platform/ballerina-library/issues) or in the [Discord server](https://discord.gg/ballerinalang). Based on the outcome of the discussion, the specification and implementation can be updated. Community contributions are also encouraged. If you notice an implementation that deviates from the specification, please raise an issue.

## Contents

1. [Overview](#1-overview)
2. [Data Types](#2-data-types)
   - 2.1. [Data](#21-data)
   - 2.2. [CellRange](#22-cellrange)
   - 2.3. [Row](#23-row)
3. [Configurations](#3-configurations)
   - 3.1. [ParseOptions](#31-parseoptions)
   - 3.2. [RowReadOptions](#32-rowreadoptions)
   - 3.3. [WriteOptions](#33-writeoptions)
   - 3.4. [RowWriteOptions](#34-rowwriteoptions)
   - 3.5. [FormulaMode](#35-formulamode)
   - 3.6. [FailSafeOptions](#36-failsafeoptions)
4. [Annotations](#4-annotations)
   - 4.1. [@xlsx:Name](#41-xlsxname)
5. [Simple API](#5-simple-api)
   - 5.1. [parse](#51-parse)
   - 5.2. [write](#52-write)
   - 5.3. [parseTable](#53-parsetable)
   - 5.4. [writeTable](#54-writetable)
   - 5.5. [parseAsStream](#55-parseasstream-reserved)
6. [Workbook API](#6-workbook-api)
   - 6.1. [Factory Functions](#61-factory-functions)
   - 6.2. [Workbook Class](#62-workbook-class)
7. [Sheet API](#7-sheet-api)
8. [Table API](#8-table-api)
9. [Error Types](#9-error-types)
10. [Samples](#10-samples)
    - 10.1. [Parse to Records](#101-parse-to-records)
    - 10.2. [Write from Records](#102-write-from-records)
    - 10.3. [Header Mapping with @xlsx:Name](#103-header-mapping-with-xlsxname)
    - 10.4. [Multi-Sheet Workbook Operations](#104-multi-sheet-workbook-operations)
    - 10.5. [Excel Tables](#105-excel-tables)
    - 10.6. [Fail-Safe Parsing](#106-fail-safe-parsing)

## 1. Overview

The `ballerinax/xlsx` module provides:

- **Type-safe XLSX parsing**: Read XLSX files into `record{}[]`, `map<anydata>[]`, or `string[][]`.
- **XLSX writing**: Write records, maps, or 2D arrays to XLSX files.
- **Workbook/Sheet API**: Multi-sheet operations with explicit lifecycle (`openFile`, `createFile`, `createWorkbook`, `save`, `saveAs`, `close`).
- **Excel Tables (ListObjects)**: Read, write, create, and modify named tables with totals row support.
- **Header-to-field mapping**: Bidirectional `@xlsx:Name` annotation for non-matching headers.
- **Used-range detection**: Excludes formatted-but-empty "ghost rows" automatically.
- **Formula handling**: Read cached values or formula text.
- **Fail-safe error handling**: Continue parsing on row-level errors with console and/or file logging.
- **Constraint validation**: Integrates with the `ballerina/constraint` module.
- **Data projection**: Lenient or strict mapping between Excel columns and record fields.

All processing is performed locally with no external service dependencies. The module uses Apache POI for XLSX file handling.

## 2. Data Types

### 2.1. Data

Union type for read/write operations.

```ballerina
public type Data anydata[][]|record {}[]|map<anydata>[];
```

| Variant | Use case |
|---|---|
| `anydata[][]` | Raw 2D array of cell values |
| `record{}[]` | Strongly-typed records; field names become column headers |
| `map<anydata>[]` | Flexible key-value rows; map keys become column headers |

### 2.2. CellRange

Represents a rectangular cell region with 0-based indices.

| Field | Type | Description |
|---|---|---|
| `firstRowIndex` | `int` | Index of the first row (0-based) |
| `lastRowIndex` | `int` | Index of the last row (0-based) |
| `firstColumnIndex` | `int` | Index of the first column (0-based) |
| `lastColumnIndex` | `int` | Index of the last column (0-based) |

Row 0 corresponds to Excel row 1, column 0 to column A.

### 2.3. Row

Wrapper type for preserving original row positions during round-trip operations. When the target record type spreads `*xlsx:Row`, parsing returns one entry per row in the used range — including empty rows as `value = null` with their `rowIndex` preserved. On write, rows are placed back at their original positions.

```ballerina
type PersonRow record {|
    *xlsx:Row;
    Person? value;
|};
```

| Field | Type | Description |
|---|---|---|
| `rowIndex` | `int` | 0-based row index relative to data start row |

## 3. Configurations

### 3.1. ParseOptions

Controls parsing behavior. Passed as the third argument to `parse()` and `parseTable()`.

| Field | Type | Default | Description |
|---|---|---|---|
| `headerRowIndex` | `int?` | `0` | Row containing column headers (0-based). Set to `null` for header-less sheets — columns become `col0`, `col1`, … |
| `dataStartRowIndex` | `int?` | unset | Row where data starts. Defaults to `headerRowIndex + 1` (or `0` when `headerRowIndex` is `null`). |
| `rowCount` | `int?` | `nil` | Maximum number of data rows to read. `null` reads all. |
| `formulaMode` | `FormulaMode` | `CACHED` | How to handle formula cells. |
| `enableConstraintValidation` | `boolean` | `true` | Validate parsed records against `@constraint` annotations. |
| `caseInsensitiveHeaders` | `boolean` | `false` | Match Excel headers to record fields case-insensitively. |
| `allowDataProjection` | record \| `false` | `{}` | Lenient mode (default) or strict (`false`). See sub-options below. |
| `failSafe` | `FailSafeOptions?` | unset | Continue on row-level errors with logging. See [3.6](#36-failsafeoptions). |

**`allowDataProjection` sub-options** (when not `false`):

| Field | Type | Default | Description |
|---|---|---|---|
| `nilAsOptionalField` | `boolean` | `false` | Treat nil cell values as field absence for optional fields. |
| `absentAsNilableType` | `boolean` | `false` | Allow missing columns for nilable/optional fields. |

### 3.2. RowReadOptions

Used by `Sheet.getRows()`, `Sheet.getRow()`, `Table.getRows()`, and `Table.getRow()`. Same fields as `ParseOptions`.

### 3.3. WriteOptions

Controls write behavior. Spread (`*WriteOptions`) on `write()` and `writeTable()`.

| Field | Type | Default | Description |
|---|---|---|---|
| `sheetName` | `string` | `"Sheet1"` | Name of the sheet to create. |
| `writeHeaders` | `boolean` | `true` | Write record field names (or map keys) as a header row. |
| `startRowIndex` | `int` | `0` | First row to write into (0-based). |

### 3.4. RowWriteOptions

Used by `Sheet.putRows()` and `Table.putRows()`.

| Field | Type | Default | Description |
|---|---|---|---|
| `writeHeaders` | `boolean` | `true` | Write headers (for record data). |
| `startRowIndex` | `int` | `0` | First row to write into (0-based). |

### 3.5. FormulaMode

```ballerina
public enum FormulaMode {
    CACHED,
    TEXT
}
```

| Value | Behavior |
|---|---|
| `CACHED` | Return the last calculated value stored in the file (default). |
| `TEXT` | Return the formula string itself (e.g., `"=SUM(A1:A10)"`). |

### 3.6. FailSafeOptions

When provided in `ParseOptions.failSafe`, parsing continues on row-level errors. Critical structural errors (missing file, corrupted workbook, missing sheet) still fail immediately.

| Field | Type | Default | Description |
|---|---|---|---|
| `enableConsoleLogs` | `boolean` | `true` | Log errors to console. |
| `includeSourceDataInConsole` | `boolean` | `false` | Include offending row data in console output. |
| `fileOutputMode` | `FileOutputMode?` | unset | Optional file-based logging. |

**`FileOutputMode`**:

| Field | Type | Default | Description |
|---|---|---|---|
| `filePath` | `string` | — | Path to the error log file. |
| `contentType` | `ErrorLogContentType` | `METADATA` | What to include in each log entry. |
| `fileWriteOption` | `FileWriteOption` | `APPEND` | `APPEND` or `OVERWRITE`. |

**`ErrorLogContentType`** enum:

| Value | Output |
|---|---|
| `METADATA` | `{"time":..., "location":{"row":N,"column":N}, "message":...}` |
| `RAW` | `["value1","value2",...]` (the offending row) |
| `RAW_AND_METADATA` | Both metadata and raw row. |

**`FileWriteOption`** enum:

| Value | Behavior |
|---|---|
| `APPEND` | Append to existing log file (default). |
| `OVERWRITE` | Overwrite on first error, append thereafter. |

The structured log entry shape is exposed as the `LogOutput` record (`time`, `location`, `message`, `offendingRow`).

## 4. Annotations

### 4.1. @xlsx:Name

Maps a record field to a specific Excel column header. Bidirectional — used on both parse and write.

```ballerina
public type NameConfig record {|
    string value;
|};

public const annotation NameConfig Name on record field;
```

```ballerina
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};
```

Fields without the annotation use their Ballerina field name as the Excel header.

## 5. Simple API

All functions are `isolated` and return the requested type or an `Error`.

### 5.1. parse

```ballerina
public isolated function parse(string path, string|int sheet = 0, ParseOptions options = {},
        typedesc<anydata[]> t = <>) returns t|Error
```

Parses an XLSX file into the inferred target type. The `sheet` parameter accepts a sheet name (string) or 0-based index (int).

Supported target types:
- `string[][]`
- `record{}[]`
- `map<anydata>[]`

Returns `FileNotFoundError`, `SheetNotFoundError`, `ParseError`, `TypeConversionError`, or `ConstraintValidationError` on failure.

### 5.2. write

```ballerina
public isolated function write(Data data, string path, *WriteOptions options) returns Error?
```

Writes data to a single-sheet XLSX file. Field names (records) or map keys become column headers when `writeHeaders` is `true`.

### 5.3. parseTable

```ballerina
public isolated function parseTable(string file, string tableName, ParseOptions options = {},
        typedesc<anydata[]> t = <>) returns t|Error
```

Parses data from a named Excel Table. Tables are unique across the workbook, so no sheet specifier is needed. Headers and the optional totals row are excluded from results automatically.

Returns `TableNotFoundError` if the table does not exist, plus the standard parse errors.

### 5.4. writeTable

```ballerina
public isolated function writeTable(Data data, string filePath, string tableName,
        *WriteOptions options) returns Error?
```

Writes data to an existing Excel Table. The table auto-expands when the data exceeds its current size. Returns `TableNotFoundError` if the table does not exist.

### 5.5. parseAsStream (reserved)

```ballerina
public isolated function parseAsStream(stream<byte[], error?> dataStream, string|int sheet = 0,
        ParseOptions options = {}, typedesc<anydata[]> t = <>) returns t|Error
```

API signature reserved for future SAX-based streaming support. Currently delegates to non-streaming parsing.

## 6. Workbook API

For multi-sheet operations, the Workbook API provides explicit lifecycle control.

### 6.1. Factory Functions

```ballerina
public function openFile(string path) returns Workbook|Error
public function createFile(string path) returns Workbook|Error
public function createWorkbook() returns Workbook|Error
```

| Function | Behavior |
|---|---|
| `openFile(path)` | Opens an existing XLSX file. `save()` overwrites it. |
| `createFile(path)` | Creates a new workbook bound to `path`. Not written until `save()` or `saveAs()` is called. |
| `createWorkbook()` | Creates an in-memory workbook with no file association. Requires `saveAs(path)` before `save()` will work. |

### 6.2. Workbook Class

| Method | Returns | Description |
|---|---|---|
| `getSheetNames()` | `string[]` | Names of all sheets, in order. |
| `getSheetCount()` | `int` | Number of sheets. |
| `getSheet(name)` | `Sheet\|SheetNotFoundError` | Get a sheet by name. |
| `getSheetByIndex(index)` | `Sheet\|SheetNotFoundError` | Get a sheet by 0-based index. |
| `createSheet(name)` | `Sheet\|Error` | Create a new sheet. Errors if the name already exists. |
| `deleteSheet(name)` | `SheetNotFoundError?` | Delete a sheet by name. |
| `deleteSheetByIndex(index)` | `SheetNotFoundError?` | Delete a sheet by 0-based index. |
| `save()` | `Error?` | Write to the workbook's source path. Errors for in-memory workbooks. |
| `saveAs(path)` | `Error?` | Write to `path` and update the source path. |
| `close()` | `Error?` | Release resources. Always call when done. |
| `getTable(name)` | `Table\|TableNotFoundError` | Get a table by name from anywhere in the workbook. |
| `getAllTables()` | `Table[]` | All tables across all sheets. |

## 7. Sheet API

A `Sheet` represents one worksheet within a workbook.

| Method | Returns | Description |
|---|---|---|
| `getName()` | `string` | Sheet name. |
| `getUsedRange()` | `string` | Used range in A1 notation (e.g., `"A1:D50"`). |
| `getUsedCellRange()` | `CellRange?` | Used range as a structured record, or `nil` if empty. |
| `getRowCount()` | `int` | Number of rows with data. |
| `getColumnCount()` | `int` | Number of columns with data. |
| `getRows(options, t)` | `t\|Error` | Read all rows as `string[][]` or `record{}[]`. |
| `getRow(index, options, t)` | `t\|Error` | Read a single row by 0-based index relative to data start. |
| `putRows(data, options)` | `Error?` | Write rows from `string[][]` or `record{}[]`. |
| `getTable(name)` | `Table\|TableNotFoundError` | Get a table in this sheet by name. |
| `getTables()` | `Table[]` | All tables in this sheet. |
| `createTable(name, range, headers?)` | `Table\|Error` | Create a table from a range (`CellRange` or A1 string). Optional explicit headers. |
| `createTableFromData(name, data, startRowIndex?, startColumnIndex?)` | `Table\|Error` | Write data and create a table around it. |
| `deleteTable(name)` | `TableNotFoundError?` | Delete a table. The underlying data is preserved. |

The "used range" excludes ghost rows — cells with formatting but no data — so iteration bounds reflect actual content.

## 8. Table API

A `Table` represents an Excel Table (ListObject). Table names are unique across the workbook.

| Method | Returns | Description |
|---|---|---|
| `getName()` | `string` | Table name. |
| `getDisplayName()` | `string` | Display name shown in Excel's UI. |
| `getSheetName()` | `string` | Name of the sheet containing the table. |
| `getRange()` | `CellRange` | Full range including headers and totals row. |
| `getDataRange()` | `CellRange` | Data range only — excludes headers and totals. |
| `getRowCount()` | `int` | Data row count. |
| `getColumnCount()` | `int` | Column count. |
| `getHeaders()` | `string[]` | Column header names in order. |
| `getRows(options, t)` | `t\|Error` | Read data rows as `string[][]` or `record{}[]`. |
| `getRow(index, options, t)` | `t\|Error` | Read a single data row by 0-based index. |
| `putRows(data, options)` | `Error?` | Write rows. Auto-expands if data exceeds current size. |
| `hasTotalsRow()` | `boolean` | Whether a totals row is present. |
| `getTotalsRow()` | `map<anydata>\|Error` | Totals values keyed by column name. |
| `rename(newName)` | `Error?` | Rename. New name must be unique workbook-wide. |
| `resize(newRange)` | `Error?` | Resize to a new range (must include header row + at least one data row). |

## 9. Error Types

The module defines a distinct error hierarchy:

```
Error (base)
├── ParseError                    — XLSX parsing failure
├── FileNotFoundError             — File missing or unreadable
├── SheetNotFoundError            — Sheet not present in workbook
├── TypeConversionError           — Cell value cannot be converted to target type
├── ConstraintValidationError     — Record violates a @constraint annotation
├── TableNotFoundError            — Table not present
├── TableOverlapError             — Creating a table would overlap with an existing one
└── InvalidTableRangeError        — Invalid range for table creation/resize
```

All error types are `distinct` subtypes of the base `Error`.

**`ErrorDetails`** record (carried by every error):

| Field | Type | Description |
|---|---|---|
| `sheetName` | `string?` | Sheet where the error occurred. |
| `tableName` | `string?` | Table where the error occurred. |
| `cellAddress` | `string?` | Cell address (e.g., `"B5"`). |
| `rowNumber` | `int?` | Row number where the error occurred. |
| `columnNumber` | `int?` | Column number where the error occurred. |

## 10. Samples

### 10.1. Parse to Records

```ballerina
import ballerinax/xlsx;

type Employee record {|
    string name;
    int age;
    string department;
|};

public function main() returns error? {
    Employee[] employees = check xlsx:parse("employees.xlsx");
}
```

### 10.2. Write from Records

```ballerina
import ballerinax/xlsx;

type Employee record {|
    string name;
    int age;
    string department;
|};

public function main() returns error? {
    Employee[] employees = [
        {name: "John", age: 30, department: "IT"},
        {name: "Jane", age: 28, department: "HR"}
    ];
    check xlsx:write(employees, "output.xlsx", sheetName = "Employees");
}
```

### 10.3. Header Mapping with @xlsx:Name

```ballerina
import ballerinax/xlsx;

type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};

public function main() returns error? {
    Employee[] employees = check xlsx:parse("employees.xlsx");
    check xlsx:write(employees, "out.xlsx");
}
```

### 10.4. Multi-Sheet Workbook Operations

```ballerina
import ballerinax/xlsx;

public function main() returns error? {
    xlsx:Workbook wb = check xlsx:openFile("report.xlsx");
    string[] sheetNames = wb.getSheetNames();

    xlsx:Sheet sales = check wb.getSheet("Sales");
    record{}[] salesRows = check sales.getRows();

    xlsx:Sheet summary = check wb.createSheet("Summary");
    check summary.putRows(salesRows);

    check wb.save();
    check wb.close();
}
```

### 10.5. Excel Tables

```ballerina
import ballerinax/xlsx;

type Employee record {|
    string name;
    int age;
    string department;
|};

public function main() returns error? {
    Employee[] current = check xlsx:parseTable("sales.xlsx", "EmployeeTable");

    Employee[] additions = [{name: "Alice", age: 31, department: "Eng"}];
    check xlsx:writeTable(additions, "sales.xlsx", "EmployeeTable");
}
```

### 10.6. Fail-Safe Parsing

```ballerina
import ballerinax/xlsx;

type Employee record {|
    string name;
    int age;
|};

public function main() returns error? {
    Employee[] employees = check xlsx:parse("data.xlsx", 0, {
        failSafe: {
            enableConsoleLogs: true,
            fileOutputMode: {
                filePath: "./xlsx-errors.log",
                contentType: xlsx:RAW_AND_METADATA,
                fileWriteOption: xlsx:OVERWRITE
            }
        }
    });
}
```
