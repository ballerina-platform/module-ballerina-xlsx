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

import ballerina/jballerina.java;

// ============================================================================
// PRIMARY API - File-based operations (recommended for most use cases)
// ============================================================================

# Parse an XLSX file into Ballerina values.
#
# This is the recommended way to read XLSX files. It reads the specified
# sheet (first sheet by default) and converts rows to the target type.
#
# Supports parsing to:
# - `string[][]` - Raw string array
# - `record{}[]` - Array of records (with header-to-field mapping)
# - `map<CellValue?>[]` - Array of maps (keys are column headers)
#
# ```ballerina
# // Parse first sheet as records
# Employee[] employees = check xlsx:parseSheet("employees.xlsx");
#
# // Parse specific sheet by name
# Employee[] sales = check xlsx:parseSheet("report.xlsx", "Sales");
#
# // Parse specific sheet by index with options
# Employee[] data = check xlsx:parseSheet("report.xlsx", 1, {headerRowIndex: 2});
# ```
#
# + path - Path to the XLSX file
# + sheet - Sheet to read: name (string) or index (int, 0-based). Default: 0 (first sheet)
# + options - Parse options
# + t - Target row type descriptor (record, map, or string[])
# + return - Parsed data or error
public isolated function parseSheet(string path, string|int sheet = 0, ParseOptions options = {},
        typedesc<Row> t = <>) returns t[]|Error = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Write Ballerina data to an XLSX file.
#
# This is the recommended way to write XLSX files. Creates a single-sheet
# XLSX file from the provided data.
#
# Supports writing from:
# - `string[][]` - Raw string array (first row can be headers)
# - `record{}[]` - Array of records (field names become headers)
# - `map<CellValue?>[]` - Array of maps (keys become headers)
#
# ```ballerina
# Employee[] employees = [{name: "John", age: 30}];
#
# // Write to file (default sheet name "Sheet1")
# check xlsx:writeSheet(employees, "output.xlsx");
#
# // Write with an explicit sheet name
# check xlsx:writeSheet(employees, "report.xlsx", "Employees");
#
# // Write with sheet name + additional row options
# check xlsx:writeSheet(employees, "report.xlsx", "Employees", writeHeaders = false);
# ```
#
# + data - Data to write
# + path - Path to the output XLSX file
# + sheetName - Name of the sheet to create (default: "Sheet1")
# + options - Row-level write options (writeHeaders, startRowIndex)
# + return - Error if write fails
public isolated function writeSheet(Row[] data, string path, string sheetName = "Sheet1",
        *RowWriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

// ============================================================================
// TABLE API - Simple functions for Excel Tables
// ============================================================================

# Parse data from an Excel table by name.
#
# Tables are unique by name across the entire workbook, so no sheet
# specification is needed. Headers are automatically excluded from results.
#
# Supports parsing to:
# - `string[][]` - Raw string array
# - `record{}[]` - Array of records (table headers map to fields)
# - `map<CellValue?>[]` - Array of maps (keys are column headers)
#
# ```ballerina
# // Parse table as records
# Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");
#
# // Parse with options
# Employee[] data = check xlsx:parseTable("report.xlsx", "SalesTable", {
#     enableConstraintValidation: true
# });
# ```
#
# + path - Path to the XLSX file
# + tableName - Name of the table to parse
# + options - Parse options
# + t - Target row type descriptor (record, map, or string[])
# + return - Parsed data or TableNotFoundError
public isolated function parseTable(string path, string tableName, ParseOptions options = {},
        typedesc<Row> t = <>) returns t[]|Error = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Write data to an existing Excel table.
#
# Writes data to the specified table. If the data exceeds the current table
# size, the table automatically expands to accommodate the new rows.
#
# ```ballerina
# Employee[] newEmployees = [...];
# check xlsx:writeTable(newEmployees, "sales.xlsx", "EmployeeTable");
# ```
#
# + data - Data to write
# + path - Path to the XLSX file containing the table
# + tableName - Name of the table to write to
# + options - Row-level write options (writeHeaders, startRowIndex)
# + return - TableNotFoundError if table doesn't exist, or other Error
public isolated function writeTable(Row[] data, string path, string tableName,
        *RowWriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Opens an XLSX workbook from a file path.
#
# Returns an error if the path does not exist or the file is not a valid XLSX.
# To create a new file, use `new Workbook()` and then `saveAs(path)`.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:fromFile("report.xlsx");
# ```
#
# + path - Path to the XLSX file
# + return - The opened workbook, or an Error if the path is missing or the file is invalid
public isolated function fromFile(string path) returns Workbook|Error = @java:Method {
    name: "openWorkbookFromPath",
    'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
} external;

# Opens an XLSX workbook from an in-memory byte array.
#
# Returns an error if the bytes are not a valid XLSX workbook.
# The resulting workbook has no associated file; use `saveAs(path)` to persist it.
#
# ```ballerina
# byte[] payload = check io:fileReadBytes("report.xlsx");
# xlsx:Workbook wb = check xlsx:fromBytes(payload);
# ```
#
# + sourceBytes - XLSX content as a byte array
# + return - The opened workbook, or an Error if the bytes are invalid
public isolated function fromBytes(byte[] sourceBytes) returns Workbook|Error = @java:Method {
    name: "openWorkbookFromBytes",
    'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
} external;
