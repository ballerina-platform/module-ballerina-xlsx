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
# - `map<anydata>[]` - Array of maps
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
    'class: "io.ballerina.stdlib.xlsx.Native"
} external;

# Write Ballerina data to an XLSX file.
#
# This is the recommended way to write XLSX files. Creates a single-sheet
# XLSX file from the provided data.
#
# Supports writing from:
# - `string[][]` - Raw string array (first row can be headers)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
#
# ```ballerina
# Employee[] employees = [{name: "John", age: 30}];
#
# // Write to file
# check xlsx:writeSheet(employees, "output.xlsx");
#
# // Write with options
# check xlsx:writeSheet(employees, "report.xlsx", sheetName = "Employees");
# ```
#
# + data - Data to write
# + path - Path to the output XLSX file
# + options - Write options
# + return - Error if write fails
public isolated function writeSheet(Data data, string path, *WriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.stdlib.xlsx.Native"
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
# - `map<anydata>[]` - Array of maps
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
    'class: "io.ballerina.stdlib.xlsx.Native"
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
# + options - Write options
# + return - TableNotFoundError if table doesn't exist, or other Error
public isolated function writeTable(Data data, string path, string tableName,
        *WriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.stdlib.xlsx.Native"
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
public isolated function fromFile(string path) returns Workbook|Error {
    Workbook wb = check new;
    check loadFromPath(wb, path);
    return wb;
}

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
public isolated function fromBytes(byte[] sourceBytes) returns Workbook|Error {
    Workbook wb = check new;
    check loadFromBytes(wb, sourceBytes);
    return wb;
}

isolated function loadFromPath(Workbook wb, string path) returns Error? = @java:Method {
    name: "openWorkbookFromPath",
    'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
} external;

isolated function loadFromBytes(Workbook wb, byte[] sourceBytes) returns Error? = @java:Method {
    name: "openWorkbookFromBytes",
    'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
} external;
