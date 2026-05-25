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
#
# + headerRowIndex - Row containing column headers/names (0-based index).
#                    Set to `null` if the sheet has no headers - columns will be named "col0", "col1", etc.
#                    Example: If headers are in row 1 (second row), set this to 1.
# + dataStartRowIndex - Row where actual data begins (0-based index).
#                       If not specified, defaults to headerRowIndex + 1 (or 0 if headerRowIndex is null).
# + rowCount - Maximum number of data rows to read. Set to `null` to read all rows (default).
#              Example: `rowCount: 100` reads at most 100 rows starting from dataStartRowIndex.
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                When enabled, parsed records are validated against any Ballerina
#                                `@constraint` annotations defined on the record type.
#                                Note: Disable for better performance when constraints aren't needed.
# + caseInsensitiveHeaders - Whether to match headers case-insensitively (default: false).
#                            When enabled, header "Name" will match record field "name" or "NAME".
# + allowDataProjection - Data projection configuration (default: enabled/lenient mode).
#                         - Default `{}`: Lenient mode - sheet columns don't need to match all record fields
#                         - Set to `false`: Strict mode - all record fields must have matching columns
#                         Sub-options when enabled:
#                         - `nilAsOptionalField`: Treat nil cells as field absence for optional fields
#                         - `absentAsNilableType`: Allow missing columns for nilable/optional fields
# + failSafe - Fail-safe error handling configuration.
#              When set, parsing continues on row-level errors (type conversion, validation).
#              Errors are logged and problematic rows are skipped.
#              Critical errors (file not found, corrupted file) still fail immediately.
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

# Options for writing XLSX data.
#
# + sheetName - Name of the sheet to create (default: "Sheet1")
# + writeHeaders - Whether to write headers from record field names (default: true)
# + startRowIndex - Row number to start writing (0-based, default: 0)
public type WriteOptions record {|
    string sheetName = "Sheet1";
    boolean writeHeaders = true;
    int startRowIndex = 0;
|};

# Options for reading rows from a sheet.
#
# + headerRowIndex - Row containing column headers/names (0-based index).
#                    Set to `null` if the sheet has no headers - columns will be named "col0", "col1", etc.
# + dataStartRowIndex - Row where actual data begins (0-based index).
#                       If not specified, defaults to headerRowIndex + 1 (or 0 if headerRowIndex is null).
# + rowCount - Maximum number of data rows to read. Set to `null` to read all rows (default).
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                When enabled, parsed records are validated against any Ballerina
#                                `@constraint` annotations defined on the record type.
#                                Note: Disable for better performance when constraints aren't needed.
# + caseInsensitiveHeaders - Whether to match headers case-insensitively (default: false).
# + allowDataProjection - Data projection configuration (see ParseOptions for details).
# + failSafe - Fail-safe error handling configuration (see ParseOptions).
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

# Options for writing rows to a sheet.
#
# + writeHeaders - Whether to write headers (default: true)
# + startRowIndex - Row number to start writing (0-based, default: 0)
public type RowWriteOptions record {|
    boolean writeHeaders = true;
    int startRowIndex = 0;
|};

# Supported data types for XLSX read/write operations.
# - `anydata[][]` - 2D array of any data (raw cell values)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
public type Data anydata[][]|record {}[]|map<anydata>[];

// =============================================================================
// Cell Range Type
// =============================================================================

# Represents a rectangular cell range in a sheet.
#
# All indices are 0-based (matching internal representation).
# For example, row 0 is the first row (Excel row 1), column 0 is column A.
#
# + firstRowIndex - Index of the first row in the range (0-based)
# + lastRowIndex - Index of the last row in the range (0-based)
# + firstColumnIndex - Index of the first column in the range (0-based)
# + lastColumnIndex - Index of the last column in the range (0-based)
public type CellRange record {|
    int firstRowIndex;
    int lastRowIndex;
    int firstColumnIndex;
    int lastColumnIndex;
|};

// =============================================================================
// Row Wrapper Type (Position Preservation)
// =============================================================================

# Row wrapper type for preserving original row positions during round-trip operations.
#
# When parsing Excel files with empty rows, positions are normally lost. Use this type
# to preserve original row positions, which is essential for:
# - Round-trip operations where formulas reference specific rows
# - Maintaining data integrity when writing back modified data
#
# **Usage:** Spread this type into your record and add a nullable `value` field:
#
# ```ballerina
# type PersonRow record {|
#     *xlsx:Row;           // Spreads rowIndex field
#     Person? value;       // Your data type (nullable for empty rows)
# |};
#
# // Parse with position preservation
# PersonRow[] rows = check xlsx:parseSheet("data.xlsx");
# // rows[0] = { rowIndex: 0, value: { name: "Alice", age: 30 } }
# // rows[1] = { rowIndex: 1, value: null }  // empty row with position preserved
# // rows[2] = { rowIndex: 2, value: { name: "Bob", age: 25 } }
#
# // Write back - positions are restored
# check xlsx:writeSheet(rows, "output.xlsx");
# ```
#
# **Behavior:**
# - When target type has `rowIndex` field, ALL rows within used range are included
# - Empty rows have `value = null` with their original `rowIndex`
# - On write, rows are placed at their original Excel positions (using `rowIndex`)
public type Row record {|
    # 0-based row index relative to data start row.
    # This preserves the original Excel row position for round-trip operations.
    int rowIndex;
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
#
# + filePath - Path to the error log file (required)
# + contentType - What content to include in logs (default: METADATA)
# + fileWriteOption - How to handle existing log files (default: APPEND)
public type FileOutputMode record {|
    string filePath;
    ErrorLogContentType contentType = METADATA;
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
# + enableConsoleLogs - Enable logging errors to console (default: true)
# + includeSourceDataInConsole - Include offending row data in console output (default: false)
# + fileOutputMode - Optional file-based error logging configuration
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
    boolean enableConsoleLogs = true;
    boolean includeSourceDataInConsole = false;
    FileOutputMode fileOutputMode?;
|};

# Location within an XLSX file where an error occurred.
#
# + row - Row number (1-based, as displayed in Excel)
# + column - Column number (1-based)
public type Location record {|
    int row;
    int column;
|};

# Structured error log output.
#
# Represents the JSON structure written to error log files when using METADATA or RAW_AND_METADATA content types.
#
# + time - ISO 8601 timestamp when the error occurred
# + location - Row and column where the error occurred
# + message - Error message describing what went wrong
# + offendingRow - The raw row data that caused the error (only with RAW_AND_METADATA)
public type LogOutput record {|
    string time?;
    Location location?;
    string message?;
    string offendingRow?;
|};
