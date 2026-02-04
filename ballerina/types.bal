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
public type NameConfig record {|
    # The Excel column header name to map to this field.
    string value;
|};

# Annotation to specify the Excel column name for a record field.
public const annotation NameConfig Name on record field;

# Options for parsing XLSX data.
#
# + headerRow - Row containing column headers/names (0-based index).
#               Set to -1 if the sheet has no headers (first row is data).
#               Example: If headers are in row 1 (second row), set this to 1.
# + dataStartRow - Row where actual data begins (0-based index).
#                  If not specified, defaults to headerRow + 1.
#                  Example: If headerRow=0, data starts at row 1 by default.
# + includeEmptyRows - Whether to include empty rows in output (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                When enabled, parsed records are validated against any Ballerina
#                                `@constraint` annotations defined on the record type.
# + caseInsensitiveHeaders - Whether to match headers case-insensitively (default: false).
#                            When enabled, header "Name" will match record field "name" or "NAME".
# + allowDataProjection - Data projection configuration.
#                         Set to `false` to disable data projection (strict mode - all fields must match).
#                         When enabled (default `{}`):
#                         - `nilAsOptionalField`: When true, nil cell values are treated as field absence
#                           (optional fields with nil values are not set in the record)
#                         - `absentAsNilableType`: When true, columns missing from the sheet are allowed
#                           for nilable/optional record fields (set to nil)
# + failSafe - Fail-safe error handling configuration.
#              When set, parsing continues on row-level errors (type conversion, validation).
#              Errors are logged and problematic rows are skipped.
#              Critical errors (file not found, corrupted file) still fail immediately.
public type ParseOptions record {|
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
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
# + startRow - Row number to start writing (0-based, default: 0)
public type WriteOptions record {|
    string sheetName = "Sheet1";
    boolean writeHeaders = true;
    int startRow = 0;
|};

# Options for reading rows from a sheet.
#
# + headerRow - Row containing column headers/names (0-based index).
#               Set to -1 if the sheet has no headers (first row is data).
# + dataStartRow - Row where actual data begins (0-based index).
#                  If not specified, defaults to headerRow + 1.
# + includeEmptyRows - Whether to include empty rows (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                When enabled, parsed records are validated against any Ballerina
#                                `@constraint` annotations defined on the record type.
# + caseInsensitiveHeaders - Whether to match headers case-insensitively (default: false).
# + allowDataProjection - Data projection configuration (see ParseOptions for details).
# + failSafe - Fail-safe error handling configuration (see ParseOptions).
public type RowReadOptions record {|
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
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
# + startRow - Row number to start writing (0-based, default: 0)
public type RowWriteOptions record {|
    boolean writeHeaders = true;
    int startRow = 0;
|};

# Supported data types for writing to XLSX files.
# - `anydata[][]` - 2D array of any data (raw cell values)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
public type WritableData anydata[][]|record {}[]|map<anydata>[];

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
# Employee[] employees = check xlsx:parse("data.xlsx", 0, {
#     failSafe: {
#         enableConsoleLogs: true,
#         includeSourceDataInConsole: true
#     }
# });
# ```
#
# # Example: File logging only
# ```ballerina
# Employee[] employees = check xlsx:parse("data.xlsx", 0, {
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
