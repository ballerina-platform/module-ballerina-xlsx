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

# How to read cells that contain a formula.
public enum FormulaMode {
    # Use the last cached value (default)
    CACHED,
    # Return the formula text, such as "=SUM(A1:A10)"
    TEXT
}

# How a sheet write treats content already at the target.
public enum SheetWriteMode {
    # Fail instead of overwriting existing content
    FAIL_IF_EXISTS,

    # Overwrite existing content in place
    REPLACE,

    # Add new content without overwriting. Any content in the way of an insert is shifted down to make room.
    APPEND
}

# How `writeTable` / `Table.putRows` treats the table's existing data.
#
# `REPLACE` (default) replaces the data and resizes the range to fit; `APPEND` adds rows below it.
# A table always has a data region, so there is no `FAIL_IF_EXISTS`.
public enum TableWriteMode {
    // REPLACE / APPEND are documented on the enum above; they are shared with SheetWriteMode, which
    // owns the per-member docs (Ballerina allows member-level metadata on only one duplicate member).
    REPLACE,
    APPEND
}

# Maps a record field to an Excel column header when the two names differ.
# Applies on both read and write; fields without it use the field name as the header.
public type NameConfig record {|
    # The Excel column header to map to this field
    string value;
|};

# Annotation to specify the Excel column name for a record field.
public const annotation NameConfig Name on record field;

# Read options shared by every read operation; the common base of all read-option records.
public type CommonParseOptions record {|
    # How to read formula cells (default: CACHED)
    FormulaMode formulaMode = CACHED;
    # Match headers case-insensitively, so "Name" matches field `name` (default: false)
    boolean caseInsensitiveHeaders = false;
|};

# Read options for sheets, which address rows by absolute position.
public type CommonSheetParseOptions record {|
    *CommonParseOptions;
    # 0-based row holding the column headers (default: 0). Set to `null` for no headers, which
    # names columns `col0`, `col1`, and so on. Ignored when reading into `string[][]`.
    int? headerRowIndex = 0;
    # 0-based row where data begins (default: `headerRowIndex` + 1, or 0 when there are no headers)
    int dataStartRowIndex?;
|};

# Data projection for record and map reads.
#
# As `{}` (default) extra columns are ignored; as `false` every record field must have a column.
public type DataProjection record {|
    # Treat nil cells as absent for optional fields (default: false)
    boolean nilAsOptionalField = false;
    # Allow missing columns for nilable fields (default: false)
    boolean absentAsNilableType = false;
|};

# Options for bulk sheet reads (`parseSheet`, `Sheet.getRows`).
public type ParseOptions record {|
    *CommonSheetParseOptions;
    # Maximum number of data rows to read (default: all)
    int rowCount?;
    # Validate parsed records against their `@constraint` annotations (default: true)
    boolean enableConstraintValidation = true;
    # Data projection: `{}` ignores extra columns (default), `false` requires an exact match
    DataProjection|false allowDataProjection = {};
    # Skip and log row-level errors instead of failing the read
    FailSafeOptions failSafe?;
|};

# Options for reading a single sheet row (`Sheet.getRow`). Fail-fast: no `failSafe` or `rowCount`.
public type RowParseOptions record {|
    *CommonSheetParseOptions;
    # Validate the parsed record against its `@constraint` annotations (default: true)
    boolean enableConstraintValidation = true;
    # Data projection: `{}` ignores extra columns (default), `false` requires an exact match
    DataProjection|false allowDataProjection = {};
|};

# Options for reading a single column (`Sheet.getColumn`).
public type ColumnParseOptions record {|
    *CommonSheetParseOptions;
    # Maximum number of cells to read (default: all)
    int rowCount?;
|};

# Options for bulk table reads (`parseTable`, `Table.getRows`).
public type TableParseOptions record {|
    *CommonParseOptions;
    # Maximum number of data rows to read (default: all)
    int rowCount?;
    # Validate parsed records against their `@constraint` annotations (default: true)
    boolean enableConstraintValidation = true;
    # Data projection: `{}` ignores extra columns (default), `false` requires an exact match
    DataProjection|false allowDataProjection = {};
    # Skip and log row-level errors instead of failing the read
    FailSafeOptions failSafe?;
|};

# Options for reading a single table row (`Table.getRow`). Fail-fast: no `failSafe` or `rowCount`.
public type TableRowParseOptions record {|
    *CommonParseOptions;
    # Validate the parsed record against its `@constraint` annotations (default: true)
    boolean enableConstraintValidation = true;
    # Data projection: `{}` ignores extra columns (default), `false` requires an exact match
    DataProjection|false allowDataProjection = {};
|};

# Options for `Sheet.putRows`.
public type WriteOptions record {|
    # Write a header row from field names or map keys (default: true; ignored for `string[][]`)
    boolean writeHeaders = true;
    # 0-based target row. Omitted, uses the mode's natural point: the end of the data for
    # `APPEND`, row 0 otherwise.
    int startRowIndex?;
    # How to treat existing content at the target (default: `APPEND`)
    SheetWriteMode sheetWriteMode = APPEND;
|};

# Options for `writeSheet`.
public type SheetWriteOptions record {|
    # Write a header row from field names or map keys (default: true; ignored for `string[][]`)
    boolean writeHeaders = true;
    # 0-based row where a fresh write starts (default: 0). Used by `FAIL_IF_EXISTS` and `REPLACE`;
    # ignored by `APPEND`, which writes below the existing data.
    int startRowIndex = 0;
    # How to treat the sheet when it already exists (default: `FAIL_IF_EXISTS`). `REPLACE` drops
    # and recreates the sheet, discarding its tables and formatting.
    SheetWriteMode sheetWriteMode = FAIL_IF_EXISTS;
|};

# Options for writing a single row (`Sheet.setRow`).
public type RowWriteOptions record {|
    # 0-based row holding the headers, used to align a record or map by name (default: 0).
    # Ignored for `string[]` data, which is written positionally.
    int headerRowIndex = 0;
    # How to treat existing content at the target row (default: `REPLACE`)
    SheetWriteMode sheetWriteMode = REPLACE;
|};

# Options for `writeTable` / `Table.putRows`.
public type TableWriteOptions record {|
    # How to treat the table's existing data (default: `REPLACE`)
    TableWriteMode tableWriteMode = REPLACE;
    # For `APPEND`, the 0-based data-row index to insert at; omitted, appends at the end.
    # Ignored by `REPLACE`.
    int insertAt?;
|};

# A single row: either a `map<CellValue>` keyed by column header, or a `string[]` of cell text
# in column order. A typed record also binds when every field type is a subtype of `CellValue`.
public type Row map<CellValue> | string[];

# An XLSX cell value: a string, number (int/float/decimal), boolean, date/time, or `()` for a blank cell.
public type CellValue string|int|float|decimal|boolean|time:Date|time:Civil|time:TimeOfDay|();

// =============================================================================
// Cell Range Type
// =============================================================================

# A rectangular cell range in a sheet, with all indices 0-based.
public type CellRange record {|
    # First row index (0-based)
    int firstRowIndex;
    # Last row index (0-based)
    int lastRowIndex;
    # First column index (0-based)
    int firstColumnIndex;
    # Last column index (0-based)
    int lastColumnIndex;
|};

// =============================================================================
// Fail-Safe Error Handling Types
// =============================================================================

# What to include when logging parsing errors.
public enum ErrorLogContentType {
    # Log only the metadata: timestamp, location, and message
    METADATA,
    # Log only the offending row data
    RAW,
    # Log both the row data and the metadata
    RAW_AND_METADATA
}

# How to write the error log file (default: `APPEND`, which adds to the existing file).
public enum FileWriteOption {
    // `APPEND` is a shared enum member whose docs are owned by `SheetWriteMode`/`TableWriteMode`
    // (Ballerina allows member-level metadata on only one occurrence), so its meaning is noted
    // in this enum doc above.
    APPEND,
    # Overwrite the log on the first error, then append the rest
    OVERWRITE
}

# File-based error logging, written to when set in `FailSafeOptions`.
public type FileOutputMode record {|
    # Path to the error log file (required)
    string filePath;
    # What to include in each log entry (default: METADATA)
    ErrorLogContentType contentType = METADATA;
    # How to write the log file (default: APPEND)
    FileWriteOption fileWriteOption = APPEND;
|};

# Fail-safe parsing: skip and log row-level errors instead of failing the read.
# Structural errors (corrupted file, missing sheet, header errors) always fail immediately.
public type FailSafeOptions record {|
    # Log errors to the console (default: true)
    boolean enableConsoleLogs = true;
    # Include the offending row data in console output (default: false)
    boolean includeSourceDataInConsole = false;
    # Also write errors to a file
    FileOutputMode fileOutputMode?;
|};

# Location within an XLSX file where an error occurred.
public type Location record {|
    # Row number (1-based, as displayed in Excel)
    int row;
    # Column number (1-based)
    int column;
|};

# The structured JSON written to error log files.
public type LogOutput record {|
    # ISO 8601 timestamp of the error
    string time?;
    # Row and column where the error occurred
    Location location?;
    # The error message
    string message?;
    # The offending row data (only with RAW_AND_METADATA)
    string offendingRow?;
|};
