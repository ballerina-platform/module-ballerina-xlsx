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

# Parse a sheet from an XLSX file into records, maps, or a string grid.
#
# ```ballerina
# Employee[] employees = check xlsx:parseSheet("employees.xlsx");
# ```
#
# + path - Path to the XLSX file
# + sheet - Sheet name or 0-based index (default: 0, the first sheet)
# + options - Parse options
# + t - Target row type
# + return - Parsed rows or an error
public isolated function parseSheet(string path, string|int sheet = 0, ParseOptions options = {},
        typedesc<Row> t = <>) returns t[]|Error = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Write rows to a sheet in an XLSX file, creating the file if it does not exist.
#
# Only the named sheet is affected; other sheets, their tables, and formulas are preserved.
# By default the write fails if the sheet already exists.
#
# ```ballerina
# Employee[] employees = [{name: "John", age: 30}];
# check xlsx:writeSheet(employees, "staff.xlsx", "Employees");
# ```
#
# + data - Rows to write (records, maps, or string arrays)
# + path - Path to the XLSX file
# + sheetName - Target sheet name (default: "Sheet1")
# + options - Write options
# + return - An error if the write fails, or if the sheet exists under FAIL_IF_EXISTS
public isolated function writeSheet(Row[] data, string path, string sheetName = "Sheet1",
        *SheetWriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

// ============================================================================
// TABLE API - Simple functions for Excel Tables
// ============================================================================

# Parse an Excel table by name into records, maps, or a string grid.
#
# Table names are unique across the workbook, so no sheet is needed. Headers and any totals
# row are excluded.
#
# ```ballerina
# Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");
# ```
#
# + path - Path to the XLSX file
# + tableName - Name of the table to parse
# + options - Table parse options
# + t - Target row type
# + return - Parsed rows or an error such as `TableNotFoundError`
public isolated function parseTable(string path, string tableName, TableParseOptions options = {},
        typedesc<Row> t = <>) returns t[]|Error = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Write rows to an existing Excel table, resizing its data range to fit.
#
# By default the table's data is replaced; `tableWriteMode = APPEND` adds rows below it instead.
#
# ```ballerina
# check xlsx:writeTable(employees, "sales.xlsx", "EmployeeTable");
# ```
#
# + data - Rows to write (records, maps, or string arrays)
# + path - Path to the XLSX file containing the table
# + tableName - Name of the table to write to
# + options - Table write options
# + return - A `TableNotFoundError`, a `TableOverlapError` if a resize collides, or another error
public isolated function writeTable(Row[] data, string path, string tableName,
        *TableWriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.lib.xlsx.Native"
} external;

# Open an XLSX workbook from a file path.
#
# To create a new file, use `new` and then `saveAs(path)`.
#
# + path - Path to the XLSX file
# + return - The opened workbook, or an error if the path is missing or the file is invalid
public isolated function fromFile(string path) returns Workbook|Error = @java:Method {
    name: "openWorkbookFromPath",
    'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
} external;

# Open an XLSX workbook from an in-memory byte array.
#
# The workbook has no associated file; use `saveAs(path)` to persist it.
#
# + sourceBytes - XLSX content as a byte array
# + return - The opened workbook, or an error if the bytes are invalid
public isolated function fromBytes(byte[] sourceBytes) returns Workbook|Error = @java:Method {
    name: "openWorkbookFromBytes",
    'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
} external;
