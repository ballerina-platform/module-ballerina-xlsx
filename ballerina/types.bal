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

# Options for parsing XLSX data.
public type ParseOptions record {|
    # Row containing column headers/names (0-based index).
    # Set to `null` if the sheet has no headers - columns will be named "col0", "col1", etc.
    # Example: If headers are in row 1 (second row), set this to 1.
    int? headerRowIndex = 0;
    # Row where actual data begins (0-based index).
    # If not specified, defaults to headerRowIndex + 1 (or 0 if headerRowIndex is null).
    int dataStartRowIndex?;
    # Maximum number of data rows to read. Set to `null` to read all rows (default).
    # Example: `rowCount: 100` reads at most 100 rows starting from dataStartRowIndex.
    int? rowCount = ();
    # How to handle formula cells (default: CACHED)
    FormulaMode formulaMode = CACHED;
    # Whether to validate type constraints (default: true).
    # When enabled, parsed records are validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    # Note: Disable for better performance when constraints aren't needed.
    boolean enableConstraintValidation = true;
    # Whether to match headers case-insensitively (default: false).
    # When enabled, header "Name" will match record field "name" or "NAME".
    boolean caseInsensitiveHeaders = false;
    # Data projection configuration (default: enabled/lenient mode).
    # - Default `{}`: Lenient mode - sheet columns don't need to match all record fields
    # - Set to `false`: Strict mode - all record fields must have matching columns
    record {|
        # Treat nil cells as field absence for optional fields
        boolean nilAsOptionalField = false;
        # Allow missing columns for nilable/optional fields
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
    # Fail-safe error handling configuration.
    # When set, parsing continues on row-level errors (type conversion, validation).
    # Errors are logged and problematic rows are skipped.
    # Critical errors (file not found, corrupted file) still fail immediately.
    FailSafeOptions failSafe?;
|};

# Options for writing XLSX data.
public type WriteOptions record {|
    # Name of the sheet to create (default: "Sheet1")
    string sheetName = "Sheet1";
    # Whether to write headers from record field names (default: true)
    boolean writeHeaders = true;
    # Row number to start writing (0-based, default: 0)
    int startRowIndex = 0;
|};

# Options for reading rows from a sheet.
public type RowReadOptions record {|
    # Row containing column headers/names (0-based index).
    # Set to `null` if the sheet has no headers - columns will be named "col0", "col1", etc.
    int? headerRowIndex = 0;
    # Row where actual data begins (0-based index).
    # If not specified, defaults to headerRowIndex + 1 (or 0 if headerRowIndex is null).
    int dataStartRowIndex?;
    # Maximum number of data rows to read. Set to `null` to read all rows (default).
    int? rowCount = ();
    # How to handle formula cells (default: CACHED)
    FormulaMode formulaMode = CACHED;
    # Whether to validate type constraints (default: true).
    # When enabled, parsed records are validated against any Ballerina
    # `@constraint` annotations defined on the record type.
    # Note: Disable for better performance when constraints aren't needed.
    boolean enableConstraintValidation = true;
    # Whether to match headers case-insensitively (default: false).
    boolean caseInsensitiveHeaders = false;
    # Data projection configuration (see ParseOptions for details).
    record {|
        # Treat nil cells as field absence for optional fields
        boolean nilAsOptionalField = false;
        # Allow missing columns for nilable/optional fields
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
    # Fail-safe error handling configuration (see ParseOptions).
    FailSafeOptions failSafe?;
|};

# Options for writing rows to a sheet.
public type RowWriteOptions record {|
    # Whether to write headers (default: true)
    boolean writeHeaders = true;
    # Row number to start writing (0-based, default: 0)
    int startRowIndex = 0;
|};

# A single row in a sheet — the atomic data unit. A row is one of three shapes:
# - `record{}` - Typed record (field names map to column headers; use `@xlsx:Name` for non-matching names)
# - `map<anydata>` - Dynamic map (keys are column headers)
# - `string[]` - Raw cell text in column order
public type Row record {} | map<anydata> | string[];

# Sheet-level data — an array of rows. Used as the return type of `parseSheet` and
# `parseTable`, and as the input type of `writeSheet` and `writeTable`.
public type Data Row[];

# Value types supported as XLSX cell content.
#
# Used as the typedesc bound for column reads (`Sheet.getColumn`). A cell in
# an XLSX file can hold a string, a number (int / float / decimal), a boolean,
# or a date/time. This union encodes that contract honestly at the type level —
# users can't ask for unsupported types like `xml` or `byte[]` from a cell and
# get a runtime error; the type system rejects it at compile time.
#
# Nilable variants (`int?`, `string?`, etc.) are supported via Ballerina's
# normal subtyping — `int?` is `int|()` and `()` is in `anydata`.
public type CellValue string|int|float|decimal|boolean|time:Date|time:Civil|time:TimeOfDay;

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
public enum FileWriteOption {
    # Append new errors to the existing log file (default).
    APPEND,
    # Overwrite the log file on first error, then append subsequent errors.
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
