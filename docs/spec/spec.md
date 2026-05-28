_Owners_: @YasanPunch \
_Reviewers_: @niveathika \
_Created_: 2026/05/02 \
_Updated_: 2026/05/24 \
_Edition_: Swan Lake

# Specification: Ballerina XLSX Module

## Introduction

This is the specification for the `xlsx` module of the [Ballerina language](https://ballerina.io/), which provides functionality for reading and writing Microsoft Excel files in the XLSX (Office Open XML) format with type-safe data binding to Ballerina records.

The `xlsx` module specification is written to describe the functionality available from version 1.0.0 onwards.

If you have any feedback or suggestions about the module, start a discussion via a [GitHub issue](https://github.com/ballerina-platform/ballerina-library/issues) or in the [Discord server](https://discord.gg/ballerinalang). Based on the outcome of the discussion, the specification and implementation can be updated. Community contributions are also encouraged. If you notice an implementation that deviates from the specification, please raise an issue.

## Contents

1. [Overview](#1-overview)
2. [Data Types](#2-data-types)
   - 2.1. [Row and Data](#21-row-and-data)
   - 2.2. [CellRange](#22-cellrange)
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
   - 5.1. [parseSheet](#51-parsesheet)
   - 5.2. [writeSheet](#52-writesheet)
   - 5.3. [parseTable](#53-parsetable)
   - 5.4. [writeTable](#54-writetable)
6. [Workbook API](#6-workbook-api)
   - 6.1. [Construction](#61-construction)
   - 6.2. [Workbook class](#62-workbook-class)
7. [Sheet API](#7-sheet-api)
8. [Table API](#8-table-api)
9. [Error Types](#9-error-types)
10. [Samples](#10-samples)
    - 10.1. [Parse to records](#101-parse-to-records)
    - 10.2. [Write from records](#102-write-from-records)
    - 10.3. [Header mapping with @xlsx:Name](#103-header-mapping-with-xlsxname)
    - 10.4. [Multi-sheet Workbook operations](#104-multi-sheet-workbook-operations)
    - 10.5. [Excel Tables](#105-excel-tables)
    - 10.6. [Bytes in, bytes out](#106-bytes-in-bytes-out)
    - 10.7. [Fail-safe error handling](#107-fail-safe-error-handling)
    - 10.8. [Date and time binding](#108-date-and-time-binding)
    - 10.9. [Large integer IDs](#109-large-integer-ids)

---

## 1. Overview

The `xlsx` module reads and writes Microsoft Excel files in the XLSX (Office Open XML) format. It provides:

- **Type-safe data binding.** XLSX rows bind to Ballerina records, maps, or string arrays. Excel column headers map to record fields automatically (overridable via `@xlsx:Name`).
- **Two API tiers.** A functional one-shot API (`parseSheet` / `writeSheet`) for simple file-based ETL, and an object-based Workbook API for multi-sheet operations, byte-array I/O, sheet/table management, and cell-level control.
- **Excel Tables (ListObjects).** Reachable two ways — tier 1 `parseTable` / `writeTable` for one-shot flows, and the `Table` class via the Workbook API for richer operations (totals row, rename, resize).
- **Fail-safe error handling.** Optional row-level error recovery with console and file logging.
- **Constraint validation.** Integrates with `ballerina/constraint` annotations on parsed records.
- **Atomic file writes.** File-based saves use a temp-file + rename pattern, so a failed write never destroys the original file.

The module uses Apache POI 5.3.0 for XLSX processing. All operations load the entire workbook into memory (DOM model); streaming is not supported in v1.0.

### v1.0 limitations

The v1.0 release deliberately defers several features. The following are **not** supported:

- **Formula authoring on write.** Strings starting with `=` are written verbatim as text, not as formula cells. There is no `Formula` wrapper type.
- **Formula re-evaluation.** `FormulaMode.CACHED` returns the last cached value as-is. There is no `EVALUATE`, `RECALCULATE`, or `PRESERVE` mode.
- **Streaming.** No row-streaming API for files larger than memory.
- **Round-trip preservation through `parseSheet`/`writeSheet`.** Tier 1 sheet functions are a data-only pipe. Formulas, formatting, charts, comments, named ranges, and other sheets are not preserved by a `parseSheet → writeSheet` cycle. (`parseTable → writeTable` preserves the surrounding workbook because it writes into an existing table; only the table's data range is overwritten.) For richer preservation, use the Workbook API and edit cells in place.
- **XLS (legacy 97-2003) format**, password-protected files, named ranges, cell styling, and range operations.

What's *included* in v1.0 that you might expect to be deferred:

- **Date / time / date-time binding** to `time:Civil` / `time:Date` / `time:TimeOfDay` (target-type-driven; ISO `string` fallback). See §10.8 for code examples.
- **Large-integer write protection**: integers with `|n| > 2^53` are written as text cells containing the exact digit string, so the data round-trips losslessly. See §10.9.
- **Excel Tables via tier 1**: `parseTable` and `writeTable` for one-shot table-by-name flows.

---

## 2. Data Types

### 2.1 Row and Data

```ballerina
# A single row in a sheet — the atomic data unit.
public type Row record {} | map<anydata> | string[];

# Sheet-level data — an array of rows.
public type Data Row[];

# Value types supported as XLSX cell content (used by `Sheet.getColumn`).
public type CellValue string|int|float|decimal|boolean
                    | time:Date|time:Civil|time:TimeOfDay;
```

`Row` is the atomic single-row type. A row can be:
- A **record** — fields named to match column headers (or via `@xlsx:Name`).
- A **`map<anydata>`** — keys are column headers; values are cell contents.
- A **`string[]`** — raw cell text in column order.

`parseSheet` takes the row shape as its target `typedesc<Row> t = <>` and returns `t[]`. `Data = Row[]` is the input type for `writeSheet` / `writeTable`. Contextual typing at the call site infers the row shape:

```ballerina
type Order record {| int id; decimal amount; |};
Order[] orders   = check xlsx:parseSheet("orders.xlsx");      // t = Order; returns Order[]
string[][] raw   = check xlsx:parseSheet("orders.xlsx");      // t = string[]; returns string[][]
map<anydata>[] m = check xlsx:parseSheet("orders.xlsx");      // t = map<anydata>; returns map<anydata>[]
```

### 2.2 CellRange

```ballerina
# A rectangular cell range with 0-based indices.
public type CellRange record {|
    int firstRowIndex;
    int lastRowIndex;
    int firstColumnIndex;
    int lastColumnIndex;
|};
```

Used by `Sheet.getUsedCellRange()`, `Sheet.createTable(name, range, headers)`, `Table.getRange()`, `Table.getDataRange()`, `Table.resize(newRange)`.

---

## 3. Configurations

### 3.1 ParseOptions

Used by `parseSheet`.

```ballerina
public type ParseOptions record {|
    int? headerRowIndex = 0;
    int dataStartRowIndex?;
    int? rowCount = ();
    FormulaMode formulaMode = CACHED;
    boolean enableConstraintValidation = true;
    boolean caseInsensitiveHeaders = false;
    record {|
        boolean nilAsOptionalField = false;
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
    FailSafeOptions failSafe?;
|};
```

| Field | Default | Meaning |
|---|---|---|
| `headerRowIndex` | `0` | 0-based row index of the header row. Set to `()` for headerless sheets — columns are exposed as `col0`, `col1`, … |
| `dataStartRowIndex` | unset | 0-based row index where data starts. Defaults to `headerRowIndex + 1` (or `0` when headerless). |
| `rowCount` | `()` | Maximum number of data rows to read. `()` reads all. |
| `formulaMode` | `CACHED` | How to handle formula cells. See [3.5](#35-formulamode). |
| `enableConstraintValidation` | `true` | When `true`, parsed records are validated against any `@constraint` annotations. |
| `caseInsensitiveHeaders` | `false` | When `true`, header `"Name"` matches record field `name` or `NAME`. |
| `allowDataProjection` | `{}` | Default `{}` enables lenient mode (extra sheet columns ignored). Set to `false` for strict mode (all record fields must have matching columns). `nilAsOptionalField` treats nil cells as field absence; `absentAsNilableType` allows missing columns for nilable/optional fields. |
| `failSafe` | unset | When set, row-level errors (type conversion, constraint validation) are logged and skipped instead of failing the parse. See [3.6](#36-failsafeoptions). |

### 3.2 RowReadOptions

Used by `Sheet.getRows`, `Sheet.getRow`, `Sheet.getColumn`, `Table.getRows`, `Table.getRow`. Same fields as `ParseOptions`.

```ballerina
public type RowReadOptions record {|
    int? headerRowIndex = 0;
    int dataStartRowIndex?;
    int? rowCount = ();
    FormulaMode formulaMode = CACHED;
    boolean enableConstraintValidation = true;
    boolean caseInsensitiveHeaders = false;
    record {|
        boolean nilAsOptionalField = false;
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
    FailSafeOptions failSafe?;
|};
```

### 3.3 WriteOptions

Used by `writeSheet` (via Ballerina spread syntax).

```ballerina
public type WriteOptions record {|
    string sheetName = "Sheet1";
    boolean writeHeaders = true;
    int startRowIndex = 0;
|};
```

| Field | Default | Meaning |
|---|---|---|
| `sheetName` | `"Sheet1"` | Name of the sheet to create. |
| `writeHeaders` | `true` | When `true`, the first row contains headers derived from record field names (or `@xlsx:Name`) or map keys. For `string[][]` input, the first row is written as-is. |
| `startRowIndex` | `0` | 0-based row index where writing begins. |

### 3.4 RowWriteOptions

Used by `Sheet.putRows`, `Sheet.setRow`, `Table.putRows`. Same fields as `WriteOptions` minus `sheetName` (the sheet is implicit in the receiver).

```ballerina
public type RowWriteOptions record {|
    boolean writeHeaders = true;
    int startRowIndex = 0;
|};
```

### 3.5 FormulaMode

```ballerina
public enum FormulaMode {
    CACHED,    # Use the last calculated/cached value (default).
    TEXT       # Return the formula string (e.g., "=SUM(A1:A10)").
}
```

- `CACHED` (default): Returns the formula cell's last cached value. The Ballerina target type must match the cached value's type.
- `TEXT`: Returns the formula expression as a string. The target field must accept `string` — otherwise a `TypeConversionError` is raised.

**Formula authoring on write is not supported in v1.0.** Strings starting with `=` are written as plain text.

### 3.6 FailSafeOptions

```ballerina
public type FailSafeOptions record {|
    boolean enableConsoleLogs = true;
    boolean includeSourceDataInConsole = false;
    FileOutputMode fileOutputMode?;
|};

public type FileOutputMode record {|
    string filePath;
    ErrorLogContentType contentType = METADATA;
    FileWriteOption fileWriteOption = APPEND;
|};

public enum ErrorLogContentType {
    METADATA,           # {"time":"...","location":{...},"message":"..."}
    RAW,                # ["value1","value2",...]
    RAW_AND_METADATA    # {"time":"...","location":{...},"offendingRow":"[...]","message":"..."}
}

public enum FileWriteOption {
    APPEND,             # Append errors to existing log file (default).
    OVERWRITE           # Overwrite log on first error, append after.
}

public type Location record {|
    int row;             # 1-based, matching Excel UI
    int column;          # 1-based
|};

public type LogOutput record {|
    string time?;
    Location location?;
    string message?;
    string offendingRow?;
|};
```

When `failSafe` is set on `ParseOptions` / `RowReadOptions`, row-level errors (`TypeConversionError`, `ConstraintValidationError`) are logged and the offending row is skipped. Structural errors still fail immediately.

---

## 4. Annotations

### 4.1 @xlsx:Name

Maps a record field to a specific Excel column header. Bidirectional — used on both read and write.

```ballerina
public type NameConfig record {|
    string value;
|};

public const annotation NameConfig Name on record field;
```

Example:

```ballerina
type Employee record {|
    @xlsx:Name {value: "First Name"}
    string firstName;
    @xlsx:Name {value: "Employee ID"}
    int id;
|};

// On read: header "First Name" binds to field firstName.
// On write: field firstName produces header "First Name".
```

---

## 5. Simple API

The simple API consists of two functions for one-shot file-based operations. Both open and close the workbook within the call.

### 5.1 parseSheet

```ballerina
public isolated function parseSheet(string path,
        string|int sheet = 0,
        ParseOptions options = {},
        typedesc<Row> t = <>)
    returns t[]|Error;
```

Reads the specified sheet from an XLSX file and binds rows to the target type inferred from the call site.

| Parameter | Default | Meaning |
|---|---|---|
| `path` | (required) | Path to the XLSX file. |
| `sheet` | `0` | Sheet selector — sheet name (`string`) or 0-based index (`int`). |
| `options` | `{}` | `ParseOptions` (see [3.1](#31-parseoptions)). |
| `t` | inferred | Target type — `Data` subtype (see [2.1](#21-row-and-data)). |

Examples:

```ballerina
// Parse first sheet as typed records.
Employee[] employees = check xlsx:parseSheet("staff.xlsx");

// Parse named sheet, raw.
string[][] rows = check xlsx:parseSheet("report.xlsx", "Q1");

// Parse with options.
Employee[] data = check xlsx:parseSheet("report.xlsx", 1,
        {headerRowIndex: 2, caseInsensitiveHeaders: true});
```

### 5.2 writeSheet

```ballerina
public isolated function writeSheet(Data data,
        string path,
        *WriteOptions options)
    returns Error?;
```

Writes data to an XLSX file, overwriting any existing file. The write is atomic — on failure, the original file is preserved.

| Parameter | Default | Meaning |
|---|---|---|
| `data` | (required) | Rows to write — `Row[]`. |
| `path` | (required) | Output file path. |
| `*options` | (defaults) | `WriteOptions` spread as named arguments (see [3.3](#33-writeoptions)). |

Examples:

```ballerina
Employee[] employees = [{name: "John", age: 30}, {name: "Jane", age: 25}];

// Simplest form.
check xlsx:writeSheet(employees, "out.xlsx");

// With options as named arguments (Ballerina spread syntax).
check xlsx:writeSheet(employees, "out.xlsx", sheetName = "Staff", writeHeaders = true);
```

**Tier 1 is data-only.** A `parseSheet → writeSheet` cycle does not preserve formulas, formatting, comments, other sheets, charts, named ranges, or Excel Tables. Use the Workbook API for any of those.

### 5.3 parseTable

```ballerina
public isolated function parseTable(string path,
        string tableName,
        ParseOptions options = {},
        typedesc<Row> t = <>)
    returns t[]|Error;
```

Reads from an Excel Table (ListObject) by name. Tables are unique by name across the entire workbook, so no sheet specifier is needed. Headers are taken from the table's own header row.

| Parameter | Default | Meaning |
|---|---|---|
| `path` | (required) | Path to the XLSX file. |
| `tableName` | (required) | Name of the table. Raises `TableNotFoundError` if no matching table exists in any sheet. |
| `options` | `{}` | `ParseOptions`. `headerRowIndex` and `dataStartRowIndex` are ignored — the table's own range defines these. All other fields (`formulaMode`, `enableConstraintValidation`, `caseInsensitiveHeaders`, `allowDataProjection`, `failSafe`) apply normally. |
| `t` | inferred | Target type — `Data` subtype. |

Example:

```ballerina
type Sale record {| string product; int quantity; decimal price; |};
Sale[] sales = check xlsx:parseTable("sales.xlsx", "SalesTable");
```

### 5.4 writeTable

```ballerina
public isolated function writeTable(Data data,
        string path,
        string tableName,
        *WriteOptions options)
    returns Error?;
```

Writes data to an existing Excel Table, auto-expanding the table's range if the data exceeds the current data range. The surrounding workbook (other sheets, named ranges, charts, formulas in unaffected cells) is preserved — only the cells inside the target table's data range are overwritten. The write is atomic.

| Parameter | Default | Meaning |
|---|---|---|
| `data` | (required) | Rows to write — `Row[]`. |
| `path` | (required) | Path to the XLSX file containing the table. |
| `tableName` | (required) | Name of the table to write into. Raises `TableNotFoundError` if no matching table exists. |
| `*options` | (defaults) | `WriteOptions` spread as named arguments. `sheetName`, `writeHeaders`, and `startRowIndex` are no-ops for `writeTable` — the table's location, headers, and data range are determined by the table itself. |

Example:

```ballerina
Sale[] sales = [{product: "Widget", quantity: 100, price: 9.99d}];
check xlsx:writeTable(sales, "sales.xlsx", "SalesTable");
```

---

## 6. Workbook API

The Workbook API exposes a stateful workbook object with explicit lifecycle.

### 6.1 Construction

Empty workbooks are constructed with `new`. To open an existing workbook from
disk or memory, use the module-level factory functions `xlsx:fromFile(path)`
and `xlsx:fromBytes(bytes)`:

```ballerina
xlsx:Workbook wb1 = check new;                              # empty in-memory
xlsx:Workbook wb2 = check xlsx:fromFile("report.xlsx");     # open existing file
xlsx:Workbook wb3 = check xlsx:fromBytes(sourceBytes);      # open from bytes
```

| Form | Semantics |
|---|---|
| `new` | Empty in-memory workbook. No source path bound. `save()` errors; `saveAs(path)` is required to persist. |
| `xlsx:fromFile(string path)` | Opens an existing file. Errors with `FileNotFoundError` if the path does not exist. To create a new file with a specific name, use `new` followed by `saveAs(path)`. |
| `xlsx:fromBytes(byte[] bytes)` | Opens the workbook represented by the byte array. No source path; `saveAs(path)` is required for file-based persistence. |

### 6.2 Workbook class

```ballerina
public isolated class Workbook {
    # Sheet access
    public isolated function getSheetNames() returns string[];
    public isolated function getSheetCount() returns int;
    public isolated function hasSheet(string name) returns boolean;
    public isolated function getSheet(string|int sheet) returns Sheet|SheetNotFoundError;
    public isolated function createSheet(string name) returns Sheet|Error;
    public isolated function deleteSheet(string|int sheet) returns Error?;

    # Table access (tables are unique by name across the workbook)
    public isolated function getTable(string name) returns Table|TableNotFoundError;
    public isolated function getAllTables() returns Table[]|Error;

    # Lifecycle
    public isolated function save() returns Error?;                   # overwrites source path; error if in-memory
    public isolated function saveAs(string path) returns Error?;      # writes to path; rebinds source path
    public isolated function toBytes() returns byte[]|Error;          # serialises workbook as XLSX bytes
    public isolated function close() returns Error?;                  # releases POI resources
}
```

**`save()` vs `saveAs()`:** `save()` overwrites the source path bound at construction (or by a previous `saveAs`). It errors for in-memory workbooks with no source path. `saveAs(path)` always writes to the given path and binds the workbook to that path so subsequent `save()` calls write there.

Both writes are atomic — temp file in the same directory + atomic rename. A failed write never destroys the original file.

**`toBytes()`** serializes the current workbook state as XLSX bytes. Useful for uploading via HTTP / SFTP without going through disk.

**`close()`** is required for resource hygiene. A phantom-reference cleanup thread catches workbooks that escape without `close()`, but explicit close is the contract.

---

## 7. Sheet API

```ballerina
public isolated class Sheet {
    # Identity and dimensions
    public isolated function getName() returns string;
    public isolated function getUsedRange() returns string;                                  # A1 notation, e.g., "A1:D50"
    public isolated function getUsedCellRange() returns CellRange?;                          # 0-based indices; nil if empty
    public isolated function getRowCount() returns int;
    public isolated function getColumnCount() returns int;

    # Row reads
    public isolated function getRows(RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error;
    public isolated function getRow(int index, RowReadOptions options = {},
            typedesc<Row> t = <>) returns t|Error;
    public isolated function getColumn(string|int columnRef, RowReadOptions options = {},
            typedesc<CellValue> t = <>) returns t[]|Error;
    public isolated function getCell(int rowIndex, int columnIndex) returns anydata|Error;

    # Row writes
    public isolated function putRows(Data data, *RowWriteOptions options) returns Error?;
    public isolated function setRow(int rowIndex, Row data, *RowWriteOptions options)
            returns Error?;
    public isolated function setColumn(string|int columnRef, anydata[] data) returns Error?;
    public isolated function setCell(int rowIndex, int columnIndex, anydata value)
            returns Error?;
    public isolated function setCellByAddress(string cellAddress, anydata value)
            returns Error?;                                                         # A1 notation

    # Sheet management
    public isolated function deleteRow(int index) returns Error?;                            # shifts subsequent rows up
    public isolated function rename(string newName) returns Error?;

    # Table access
    public isolated function getTable(string name) returns Table|TableNotFoundError;
    public isolated function getTables() returns Table[]|Error;
    public isolated function createTable(string name, CellRange|string range,
            string[]? headers = ()) returns Table|Error;
    public isolated function createTableFromData(string name, Data data,
            int startRowIndex = 0, int startColumnIndex = 0) returns Table|Error;
    public isolated function deleteTable(string name) returns TableNotFoundError?;
}
```

Notes:
- `Sheet.getCell` returns `anydata` because the type depends on the cell's content. Callers narrow as needed.
- `Sheet.getColumn` accepts a column reference as either a header name (`string`) or a 0-based index (`int`).
- `Sheet.deleteRow(index)` removes the row and shifts subsequent rows up by one to preserve dense indexing.

---

## 8. Table API

```ballerina
public isolated class Table {
    # Identity
    public isolated function getName() returns string;
    public isolated function getDisplayName() returns string;
    public isolated function getSheetName() returns string;

    # Range and dimensions (0-based indices)
    public isolated function getRange() returns CellRange;          # full table including headers/totals
    public isolated function getDataRange() returns CellRange;      # data rows only
    public isolated function getRowCount() returns int;             # data rows only
    public isolated function getColumnCount() returns int;

    # Headers and data
    public isolated function getHeaders() returns string[];
    public isolated function getRows(RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error;
    public isolated function getRow(int index, RowReadOptions options = {},
            typedesc<Row> t = <>) returns t|Error;
    public isolated function putRows(Data data, *RowWriteOptions options) returns Error?;   # auto-expands the table

    # Totals row
    public isolated function hasTotalsRow() returns boolean;
    public isolated function getTotalsRow() returns map<anydata>|Error;

    # Modification
    public isolated function rename(string newName) returns Error?;
    public isolated function resize(CellRange newRange) returns Error?;
}
```

Tables are obtained from `Workbook.getTable(name)`, `Workbook.getAllTables()`, `Sheet.getTable(name)`, `Sheet.getTables()`, `Sheet.createTable(...)`, or `Sheet.createTableFromData(...)`. Table names are unique across the entire workbook.

`Table.putRows` automatically expands the underlying `XSSFTable` to fit the incoming data if it exceeds the current data range.

---

## 9. Error Types

```ballerina
public type Error distinct error<ErrorDetails>;
public type ParseError distinct Error;
public type FileNotFoundError distinct Error;
public type SheetNotFoundError distinct Error;
public type TypeConversionError distinct Error;
public type ConstraintValidationError distinct Error;
public type TableNotFoundError distinct Error;
public type TableOverlapError distinct Error;
public type InvalidTableRangeError distinct Error;

public type ErrorDetails record {|
    string sheetName?;
    string tableName?;
    string cellAddress?;     # A1 notation, e.g., "B5"
    int rowNumber?;          # 1-based, matching Excel UI
    int columnNumber?;       # 1-based, matching Excel UI
|};
```

Semantics:
- **Structural errors** (`ParseError`, `FileNotFoundError`, `SheetNotFoundError`, `TableNotFoundError`, `TableOverlapError`, `InvalidTableRangeError`) fail immediately, regardless of `failSafe`.
- **Row-level errors** (`TypeConversionError`, `ConstraintValidationError`) fail immediately by default; with `failSafe` set, the offending row is logged and skipped.

---

## 10. Samples

### 10.1 Parse to records

```ballerina
import ballerina/xlsx;

type Employee record {|
    string name;
    int age;
    decimal salary;
|};

public function main() returns error? {
    Employee[] employees = check xlsx:parseSheet("staff.xlsx");
    foreach Employee emp in employees {
        // process each employee
    }
}
```

### 10.2 Write from records

```ballerina
import ballerina/xlsx;

type Employee record {|
    string name;
    int age;
    decimal salary;
|};

public function main() returns error? {
    Employee[] employees = [
        {name: "Alice", age: 30, salary: 75000d},
        {name: "Bob", age: 25, salary: 60000d}
    ];
    check xlsx:writeSheet(employees, "staff.xlsx", sheetName = "Employees");
}
```

### 10.3 Header mapping with @xlsx:Name

```ballerina
import ballerina/xlsx;

type Employee record {|
    @xlsx:Name {value: "First Name"}
    string firstName;
    @xlsx:Name {value: "Employee ID"}
    int id;
|};

public function main() returns error? {
    // Excel columns: "First Name" | "Employee ID"
    Employee[] employees = check xlsx:parseSheet("staff.xlsx");

    // Round-trips: write produces the same "First Name" / "Employee ID" headers.
    check xlsx:writeSheet(employees, "staff-out.xlsx");
}
```

### 10.4 Multi-sheet Workbook operations

```ballerina
import ballerina/xlsx;

type Sale record {| string product; int quantity; decimal price; |};

public function main() returns error? {
    xlsx:Workbook wb = check xlsx:fromFile("sales.xlsx");

    // Read from one sheet, modify, write to another.
    xlsx:Sheet rawSheet = check wb.getSheet("Raw");
    Sale[] sales = check rawSheet.getRows();

    Sale[] highValue = from Sale s in sales where s.price > 100d select s;

    xlsx:Sheet summarySheet = check wb.createSheet("HighValue");
    check summarySheet.putRows(highValue);

    check wb.save();
    check wb.close();
}
```

### 10.5 Excel Tables

Two paths, depending on whether you need the broader workbook context.

**Tier 1 — one-shot table read/write:**

```ballerina
import ballerina/xlsx;

type Employee record {| string name; int age; |};

public function main() returns error? {
    // Read all rows of a named table.
    Employee[] employees = check xlsx:parseTable("data.xlsx", "EmployeeTable");

    // Append a row and write back. writeTable auto-expands the table.
    Employee[] withNew = [...employees, {name: "Charlie", age: 35}];
    check xlsx:writeTable(withNew, "data.xlsx", "EmployeeTable");
}
```

**Workbook API — when you need totals row, rename/resize, or coordination with other operations:**

```ballerina
import ballerina/xlsx;

type Employee record {| string name; int age; |};

public function main() returns error? {
    xlsx:Workbook wb = check xlsx:fromFile("data.xlsx");

    xlsx:Table empTable = check wb.getTable("EmployeeTable");
    Employee[] employees = check empTable.getRows();

    if empTable.hasTotalsRow() {
        map<anydata> totals = check empTable.getTotalsRow();
        // ... inspect totals ...
    }

    Employee[] newEmployees = [...employees, {name: "Charlie", age: 35}];
    check empTable.putRows(newEmployees);

    check wb.save();
    check wb.close();
}
```

### 10.6 Bytes in, bytes out

```ballerina
import ballerina/ftp;
import ballerina/xlsx;

type Order record {| int id; string customer; decimal amount; |};

public function main() returns error? {
    ftp:Client sftp = check new ({host: "sftp.example.com"});

    // Pull bytes from SFTP, open as a workbook.
    byte[] inputBytes = check sftp->get("/in/orders.xlsx");
    xlsx:Workbook wb = check xlsx:fromBytes(inputBytes);

    xlsx:Sheet sheet = check wb.getSheet(0);
    Order[] orders = check sheet.getRows();

    // Enrich and write back into the same sheet.
    Order[] enriched = from Order o in orders select {...o, amount: o.amount * 1.1d};
    check sheet.putRows(enriched);

    // Serialise and upload.
    byte[] outputBytes = check wb.toBytes();
    check sftp->put("/out/orders-enriched.xlsx", outputBytes);

    check wb.close();
}
```

### 10.7 Fail-safe error handling

```ballerina
import ballerina/xlsx;

type Employee record {|
    string name;
    int age;
|};

public function main() returns error? {
    // Bad rows logged to a file; parse returns the good rows.
    Employee[] employees = check xlsx:parseSheet("messy.xlsx", 0, {
        failSafe: {
            enableConsoleLogs: true,
            includeSourceDataInConsole: true,
            fileOutputMode: {
                filePath: "./errors.log",
                contentType: RAW_AND_METADATA,
                fileWriteOption: APPEND
            }
        }
    });
    // Use the cleaned-up records.
}
```

### 10.8 Date and time binding

The binder uses the target field type to decide what shape to produce. Declare the field as `time:Civil` / `time:Date` / `time:TimeOfDay` for typed values, or as `string` for ISO 8601.

```ballerina
import ballerina/xlsx;
import ballerina/time;

type Transaction record {|
    int id;
    time:Civil timestamp;     // date-time cell → time:Civil
    time:Date settledOn;      // date-only cell → time:Date
    decimal amount;
|};

public function main() returns error? {
    Transaction[] txns = check xlsx:parseSheet("transactions.xlsx");

    // Work with the values as time records — no manual parsing.
    foreach Transaction t in txns {
        if t.settledOn.year >= 2026 {
            // ...
        }
    }

    // Writing back produces date-formatted cells, not text cells.
    check xlsx:writeSheet(txns, "transactions-out.xlsx");
}
```

If you prefer ISO strings (e.g., when the target type is dynamic), declare the fields as `string`:

```ballerina
type RawTxn record {|
    int id;
    string timestamp;         // "2026-05-24T10:30:00"
    string settledOn;         // "2026-05-24"
    decimal amount;
|};

RawTxn[] raw = check xlsx:parseSheet("transactions.xlsx");
```

### 10.9 Large integer IDs

Integers with absolute value greater than `2^53` (≈ 9 × 10^15) cannot be represented exactly as IEEE-754 doubles. v1.0 writes them as text cells with the exact digit string preserved.

```ballerina
import ballerina/xlsx;

type Order record {|
    int orderId;              // e.g., 4929187654321098765 — 19 digits
    string customer;
    decimal amount;
|};

public function main() returns error? {
    Order[] orders = [
        {orderId: 4929187654321098765, customer: "Acme", amount: 99.99d}
    ];

    // The orderId is written as a text cell. The digits round-trip exactly.
    check xlsx:writeSheet(orders, "orders.xlsx");

    Order[] readBack = check xlsx:parseSheet("orders.xlsx");
    // readBack[0].orderId == 4929187654321098765 — preserved
}
```

In Excel the affected cells appear as text (left-aligned, no numeric formatting). If you need Excel to treat the column as numeric for display purposes, declare the field as `string` in your record to make the intent explicit; v1.0 doesn't auto-format the cell back as numeric.
