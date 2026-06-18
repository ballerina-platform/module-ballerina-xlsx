// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/time;

# Formula handling mode for cells containing formulas.
public enum FormulaMode {
    # Use the last calculated/cached value (default).
    CACHED,
    # Return the formula string as text (e.g., "=SUM(A1:A10)").
    TEXT
}

# How a sheet write treats content already present at the target.
#
# Shared by `writeSheet` (target = the named sheet) and the row writers `Sheet.putRows` /
# `Sheet.setRow` (target = the rows being written). The default differs per operation —
# `FAIL_IF_EXISTS` for `writeSheet`, `APPEND` for `putRows`, `REPLACE` for `setRow`.
public enum SheetWriteMode {
    # Fail rather than touch existing content. `writeSheet` errors if the named sheet already
    # exists; `putRows` / `setRow` error if the target rows are not empty.
    FAIL_IF_EXISTS,

    # Overwrite existing content in place. `writeSheet` drops and recreates the named sheet
    # (siblings preserved; the sheet's own formatting and any table on it are lost);
    # `putRows` / `setRow` overwrite the target rows.
    REPLACE,

    # Add rows, preserving existing content by shifting it down to make room. `writeSheet`
    # appends below the sheet's data; `putRows` / `setRow` insert at the target row, or append at
    # the bottom when no explicit position is given. Record/map writes align to the existing
    # header row by column name.
    APPEND
}

# Behaviour of `writeTable` / `Table.putRows` toward the table's existing data.
#
# A table always has a data region, so there is no requirement for `FAIL_IF_EXISTS`. Plain content
# below the table is carried along with any shift; if a grow collides with another table,
# the write fails with a `TableOverlapError`.
#
# - `REPLACE` (default): replace the table's data with the given rows, resizing the data range to
#   fit exactly (grows or shrinks). Writing an empty array clears the table to a single blank row.
# - `APPEND`: add the given rows below the table's existing data (or at `insertAt`), shifting any
#   rows below down to make room. Existing data rows are preserved.
public enum TableWriteMode {
    // REPLACE / APPEND are documented on the enum above; they are shared with SheetWriteMode, which
    // owns the per-member docs (Ballerina allows member-level metadata on only one duplicate member).
    REPLACE,
    APPEND
}

# Annotation to map a record field to a specific Excel column name.
# Use this when the Excel column header doesn't match the Ballerina field name.
#
# ```ballerina
# type Employee record {
#     @xlsx:Name {value: "First Name"}
#     string firstName;
#     @xlsx:Name {value: "Employee ID"}
#     int id;
# };
# ```
#
# When reading, headers "First Name" and "Employee ID" will map to `firstName` and `id`.
# When writing, field names will produce headers "First Name" and "Employee ID".
# Fields without this annotation use their Ballerina field name as the Excel header.
public type NameConfig record {|
    # The Excel column header name to map to this field.
    string value;
|};

# Annotation to specify the Excel column name for a record field.
public const annotation NameConfig Name on record field;

# Cell- and record-binding options honoured by every read operation, regardless of source
# (sheet or table), cardinality (bulk, single row, single column), or result shape. Included
# as the common base of every read-option record via `*CommonParseOptions;`.
#
# The positional row-window fields (`headerRowIndex`, `dataStartRowIndex`) are deliberately
# *not* here: they apply only to sheet reads — a table is self-describing (its column
# definitions are the header, its area is the data range) — so they live one level down in
# `CommonSheetParseOptions`.
public type CommonParseOptions record {|
    # How to handle formula cells (default: CACHED)
    FormulaMode formulaMode = CACHED;
    # Whether to match headers case-insensitively (default: false).
    # When enabled, header "Name" will match record field "name" or "NAME".
    boolean caseInsensitiveHeaders = false;
|};

# Options shared by sheet reads, which address rows by absolute position. Adds the
# row-window fields to `CommonParseOptions`; included by `ParseOptions`, `RowParseOptions`,
# and `ColumnParseOptions`. Table reads omit these and use `TableParseOptions` /
# `TableRowParseOptions` instead.
public type CommonSheetParseOptions record {|
    *CommonParseOptions;
    # Row containing column headers/names (0-based index).
    # Set to `null` if the sheet has no headers - columns will be named "col0", "col1", etc.
    # Example: If headers are in row 1 (second row), set this to 1.
    # Note: when reading into `string[][]`, the header row is returned as data (raw mode is
    # lossless) and this field is ignored - use `dataStartRowIndex` to skip leading rows.
    int? headerRowIndex = 0;
    # Row where actual data begins (0-based index).
    # If not specified, defaults to headerRowIndex + 1 (or 0 if headerRowIndex is null).
    int dataStartRowIndex?;
|};

# Data projection configuration for record/map reads.
#
# - Default `{}`: lenient mode — columns don't need to match every record field.
# - Set to `false` (in the enclosing option's `allowDataProjection`): strict mode — every
#   record field must have a matching column.
public type DataProjection record {|
    # Treat nil cells as field absence for optional fields.
    boolean nilAsOptionalField = false;
    # Allow missing columns for nilable/optional fields.
    boolean absentAsNilableType = false;
|};

# Options for bulk sheet reads — `parseSheet`, `Sheet.getRows`.
#
# Reads the row window `[dataStartRowIndex, dataStartRowIndex + rowCount)`; with `rowCount`
# unset (`null`) it reads through the last used row.
public type ParseOptions record {|
    *CommonSheetParseOptions;
    # Maximum number of data rows to read. Set to `null` to read all rows (default).
    # Example: `rowCount: 100` reads at most 100 rows starting from dataStartRowIndex.
    int? rowCount = ();
    # Whether to validate type constraints (default: true).
    # When enabled, parsed records are validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    # Note: Disable for better performance when constraints aren't needed.
    boolean enableConstraintValidation = true;
    # Data projection configuration (default: lenient `{}`; set to `false` for strict mode).
    DataProjection|false allowDataProjection = {};
    # Fail-safe error handling configuration.
    # When set, parsing continues on row-level errors (type conversion, validation).
    # Errors are logged and problematic rows are skipped.
    # Critical errors (file not found, corrupted file) still fail immediately.
    FailSafeOptions failSafe?;
|};

# Options for reading a single sheet row — `Sheet.getRow`.
#
# Single-row reads are fail-fast, so there is no `failSafe` (skipping the only requested
# row would leave nothing to return) and no `rowCount`. Constraint validation and data
# projection still apply.
public type RowParseOptions record {|
    *CommonSheetParseOptions;
    # Whether to validate type constraints (default: true).
    # When enabled, the parsed record is validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    boolean enableConstraintValidation = true;
    # Data projection configuration (see `ParseOptions` for details).
    DataProjection|false allowDataProjection = {};
|};

# Options for reading a single column — `Sheet.getColumn`.
#
# A column read yields scalar cell values rather than records, so constraint validation,
# data projection, and fail-safe (record/bulk concerns) do not apply.
public type ColumnParseOptions record {|
    *CommonSheetParseOptions;
    # Maximum number of cells to read. Set to `null` to read all (default).
    int? rowCount = ();
|};

# Options for bulk table reads — `parseTable`, `Table.getRows`.
#
# A table is self-describing: its column definitions are the header and its area is the data
# range, so there are no positional `headerRowIndex` / `dataStartRowIndex` fields. Reads the
# first `rowCount` data rows (the header and any totals row are always excluded); with
# `rowCount` unset (`null`) it reads every data row.
public type TableParseOptions record {|
    *CommonParseOptions;
    # Maximum number of data rows to read. Set to `null` to read all rows (default).
    int? rowCount = ();
    # Whether to validate type constraints (default: true).
    # When enabled, parsed records are validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    # Note: Disable for better performance when constraints aren't needed.
    boolean enableConstraintValidation = true;
    # Data projection configuration (default: lenient `{}`; set to `false` for strict mode).
    DataProjection|false allowDataProjection = {};
    # Fail-safe error handling configuration.
    # When set, parsing continues on row-level errors (type conversion, validation).
    # Errors are logged and problematic rows are skipped.
    # Critical errors (file not found, corrupted file) still fail immediately.
    FailSafeOptions failSafe?;
|};

# Options for reading a single table row — `Table.getRow`.
#
# Self-describing like `TableParseOptions` (no positional fields), and fail-fast like
# `RowParseOptions` (no `failSafe`, no `rowCount`). Constraint validation and data
# projection still apply.
public type TableRowParseOptions record {|
    *CommonParseOptions;
    # Whether to validate type constraints (default: true).
    # When enabled, the parsed record is validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    boolean enableConstraintValidation = true;
    # Data projection configuration (see `TableParseOptions` for details).
    DataProjection|false allowDataProjection = {};
|};

# Options for `Sheet.putRows`.
#
# `startRowIndex` is the target position; left unset (`()`) it resolves to the mode's natural
# point — the bottom of the used range for `APPEND`, row 0 for `REPLACE` / `FAIL_IF_EXISTS`.
public type WriteOptions record {|
    # Whether to write a header row (default: true). Headers are derived from record field
    # names (honouring `@xlsx:Name`) or map keys. Ignored for `string[][]` input, whose
    # first row is written as-is.
    boolean writeHeaders = true;
    # Target row (0-based). Unset (`()`) = the mode's natural point: the bottom of the used range
    # for `APPEND`, row 0 for `REPLACE` / `FAIL_IF_EXISTS`. With `APPEND` an explicit value inserts
    # at that row (shifting existing rows down); with `REPLACE` it overwrites from that row.
    int? startRowIndex = ();
    # How the write treats existing content at the target (default: `APPEND` — add rows
    # non-destructively, inserting and shifting rather than overwriting).
    SheetWriteMode sheetWriteMode = APPEND;
|};

# Options for `writeSheet`.
public type SheetWriteOptions record {|
    # Whether to write a header row (default: true). Headers are derived from record field
    # names (honouring `@xlsx:Name`) or map keys. Ignored for `string[][]` input, whose
    # first row is written as-is.
    boolean writeHeaders = true;
    # Row at which a fresh write block starts (0-based, default: 0) — the header goes here and
    # data follows. Used only by `FAIL_IF_EXISTS` and `REPLACE`. Ignored by `APPEND`, which always
    # appends below the sheet's existing data and auto-detects the header from the used range.
    int startRowIndex = 0;
    # How the target sheet is treated when the file already contains it (default: `FAIL_IF_EXISTS`).
    # Sibling sheets are preserved in every mode. **`REPLACE` is destructive** — it drops and
    # recreates the named sheet, discarding any tables and formatting on it; the default
    # `FAIL_IF_EXISTS` guards against accidental overwrites.
    SheetWriteMode sheetWriteMode = FAIL_IF_EXISTS;
|};

# Options for writing a single row — `Sheet.setRow`.
#
# A single-row write targets an explicit row index. For record or map data it needs to know
# where the header row lives in order to align values to columns by name.
public type RowWriteOptions record {|
    # Row containing the column headers used to align a record/map row to columns (0-based,
    # default: 0). Ignored for `string[]` data, which is written positionally.
    int headerRowIndex = 0;
    # How the write treats existing content at the target row (default: `REPLACE` — overwrite the
    # row). `APPEND` inserts a row there (shifting existing rows down); `FAIL_IF_EXISTS` errors if
    # the target row is not empty.
    SheetWriteMode sheetWriteMode = REPLACE;
|};

# Options for `writeTable` / `Table.putRows`.
public type TableWriteOptions record {|
    # How the write treats the table's existing data (default: `REPLACE`).
    TableWriteMode tableWriteMode = REPLACE;
    # For `APPEND`, the 0-based data-row index to insert at — existing rows from there down, the
    # totals row, and any content below shift down to make room. Unset (`()`) appends at the bottom.
    # Ignored by `REPLACE`, which replaces the whole data region.
    int? insertAt = ();
|};

# A single row in a sheet — the atomic data unit. A row is one of two shapes:
# - `map<CellValue>` - Dynamic map (keys are column headers; values are cell values,
#   with `()` for a blank cell — the empty cell is a member of `CellValue`). A typed
#   record also binds when every field type is a subtype of `CellValue`; use `@xlsx:Name`
#   for field names that differ from the column headers. To absorb columns beyond the
#   declared fields, give the record a `CellValue` rest descriptor —
#   `record {| ...; CellValue...; |}`.
# - `string[]` - Raw cell text in column order
#
# The map's value type is `CellValue` (not `anydata`) so the row contract matches what
# a cell can actually hold: a target field of an unsupported type (e.g. `xml`, `byte[]`,
# a nested record) is rejected at compile time rather than failing at runtime.
public type Row map<CellValue> | string[];

# An XLSX cell value, including the empty cell.
#
# A populated cell holds a string, a number (int / float / decimal), a boolean, or a
# date/time; a blank cell is `()`. This union encodes that contract honestly at the type
# level — users can't ask for unsupported types like `xml` or `byte[]` from a cell and
# get a runtime error; the type system rejects it at compile time.
#
# Because the empty cell is a member of the union, a blank cell binds to `()` directly:
# the scalar-cell APIs (`Sheet.getCell`, `Sheet.getColumn`, the cell setters, and
# `Table.getTotalRow`) use `CellValue` without a trailing `?`.
public type CellValue string|int|float|decimal|boolean|time:Date|time:Civil|time:TimeOfDay|();

// =============================================================================
// Cell Range Type
// =============================================================================

# Represents a rectangular cell range in a sheet.
#
# All indices are 0-based (matching internal representation).
# For example, row 0 is the first row (Excel row 1), column 0 is column A.
public type CellRange record {|
    # Index of the first row in the range (0-based)
    int firstRowIndex;
    # Index of the last row in the range (0-based)
    int lastRowIndex;
    # Index of the first column in the range (0-based)
    int firstColumnIndex;
    # Index of the last column in the range (0-based)
    int lastColumnIndex;
|};

// =============================================================================
// Fail-Safe Error Handling Types
// =============================================================================

# Content types for error log output.
#
# Controls what information is included when logging parsing errors.
public enum ErrorLogContentType {
    # Log only metadata: timestamp, location (row/column), error message.
    # Output format: `{"time":"...","location":{"row":N,"column":N},"message":"..."}`
    METADATA,
    # Log only the raw offending row data as a JSON array.
    # Output format: `["value1", "value2", ...]`
    RAW,
    # Log both raw data and metadata for comprehensive debugging.
    # Output format: `{"time":"...","location":{...},"offendingRow":"[...]","message":"..."}`
    RAW_AND_METADATA
}

# File write options for error logs.
#
# Controls how the error log file is written when multiple errors occur.
#
# - `APPEND` (default): append new errors to the existing log file.
# - `OVERWRITE`: overwrite the log file on the first error, then append subsequent errors.
public enum FileWriteOption {
    // APPEND is shared with SheetWriteMode/TableWriteMode, which own its per-member doc; both
    // members are documented on the enum above (member-level metadata is allowed on only one
    // occurrence of a duplicated enum member).
    APPEND,
    OVERWRITE
}

# Configuration for file-based error logging.
#
# When provided in `FailSafeOptions`, parsing errors will be written to the specified file.
public type FileOutputMode record {|
    # Path to the error log file (required)
    string filePath;
    # What content to include in logs (default: METADATA)
    ErrorLogContentType contentType = METADATA;
    # How to handle existing log files (default: APPEND)
    FileWriteOption fileWriteOption = APPEND;
|};

# Configuration for fail-safe error handling during parsing.
#
# When enabled, parsing continues even when row-level errors occur (e.g., type conversion failures).
# Errors can be logged to console, file, or both. Rows with errors are skipped, and only
# successfully parsed rows are returned.
#
# **Note**: Critical structural errors (corrupted file, missing sheet, header errors) will
# always cause parsing to fail immediately, regardless of fail-safe configuration.
#
# # Example: Console logging only
# ```ballerina
# Employee[] employees = check xlsx:parseSheet("data.xlsx", 0, {
#     failSafe: {
#         enableConsoleLogs: true,
#         includeSourceDataInConsole: true
#     }
# });
# ```
#
# # Example: File logging only
# ```ballerina
# Employee[] employees = check xlsx:parseSheet("data.xlsx", 0, {
#     failSafe: {
#         enableConsoleLogs: false,
#         fileOutputMode: {
#             filePath: "./xlsx-errors.log",
#             contentType: RAW_AND_METADATA,
#             fileWriteOption: OVERWRITE
#         }
#     }
# });
# ```
public type FailSafeOptions record {|
    # Enable logging errors to console (default: true)
    boolean enableConsoleLogs = true;
    # Include offending row data in console output (default: false)
    boolean includeSourceDataInConsole = false;
    # Optional file-based error logging configuration
    FileOutputMode fileOutputMode?;
|};

# Location within an XLSX file where an error occurred.
public type Location record {|
    # Row number (1-based, as displayed in Excel)
    int row;
    # Column number (1-based)
    int column;
|};

# Structured error log output.
#
# Represents the JSON structure written to error log files when using METADATA or RAW_AND_METADATA content types.
public type LogOutput record {|
    # ISO 8601 timestamp when the error occurred
    string time?;
    # Row and column where the error occurred
    Location location?;
    # Error message describing what went wrong
    string message?;
    # The raw row data that caused the error (only with RAW_AND_METADATA)
    string offendingRow?;
|};
