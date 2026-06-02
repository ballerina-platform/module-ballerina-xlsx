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

# Represents an Excel Table (ListObject) within a worksheet.
#
# Tables provide structured access to data with automatic header handling,
# optional totals row support, and auto-expand capability when writing.
# Table names are unique across the entire workbook.
#
# Instances are obtained from a `Workbook` or `Sheet` via methods like
# `getTable`, `createTable`, etc.; direct construction (`new Table()`) is not supported.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:fromFile("sales.xlsx");
# xlsx:Table empTable = check wb.getTable("EmployeeTable");
#
# // Read data (headers excluded automatically)
# Employee[] employees = check empTable.getRows();
#
# // Write data (auto-expands if needed)
# check empTable.putRows(newEmployees);
#
# check wb.save();
# check wb.close();
# ```
public type Table isolated object {

    # Get the name of the table.
    #
    # Table names are unique across the entire workbook.
    #
    # + return - Table name, or an `Error` if the handle is invalid
    public isolated function getName() returns string|Error;

    # Get the display name of the table.
    #
    # The display name is what appears in Excel's UI. It may differ from the internal name.
    #
    # + return - Table display name, or an `Error` if the handle is invalid
    public isolated function getDisplayName() returns string|Error;

    # Get the name of the sheet containing this table.
    #
    # + return - Sheet name, or an `Error` if the handle is invalid
    public isolated function getSheetName() returns string|Error;

    # Get the full range of the table (including headers and total row) in A1 notation.
    #
    # + return - A1-notation range string (e.g., "A1:D10"), or an `Error` if the handle is invalid
    public isolated function getRange() returns string|Error;

    # Get the full range of the table (including headers and total row) as a `CellRange`.
    #
    # All indices are 0-based.
    #
    # + return - CellRange representing the full table area, or an `Error` if the handle is invalid
    public isolated function getCellRange() returns CellRange|Error;

    # Get the data range of the table (excluding headers and total row) in A1 notation.
    #
    # + return - A1-notation range string, or an `Error` if the handle is invalid
    public isolated function getDataRange() returns string|Error;

    # Get the data range of the table (excluding headers and total row) as a `CellRange`.
    #
    # All indices are 0-based.
    #
    # + return - CellRange representing only the data area, or an `Error` if the handle is invalid
    public isolated function getDataCellRange() returns CellRange|Error;

    # Get the number of data rows in the table.
    #
    # Returns only data rows, excluding header and totals row.
    #
    # + return - Data row count, or an `Error` if the handle is invalid
    public isolated function getRowCount() returns int|Error;

    # Get the number of columns in the table.
    #
    # + return - Column count, or an `Error` if the handle is invalid
    public isolated function getColumnCount() returns int|Error;

    # Get the column header names.
    #
    # Returns an array of header names in column order.
    #
    # + return - Array of header strings, or an `Error` if the handle is invalid
    public isolated function getHeaders() returns string[]|Error;

    # Get all data rows from the table.
    #
    # Headers and totals row are automatically excluded. Supports reading to:
    # - `string[][]` - Raw string array
    # - `record{}[]` - Array of records (headers map to fields)
    #
    # ```ballerina
    # // As string array
    # string[][] rows = check table.getRows();
    #
    # // As records
    # type Employee record {| string name; int age; |};
    # Employee[] employees = check table.getRows();
    # ```
    #
    # + options - Read options
    # + t - Target row type descriptor (record, map, or string[])
    # + return - Array of data rows or error
    public isolated function getRows(RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error;

    # Get a single data row from the table by index.
    #
    # The index is 0-based within the data range (first data row is index 0).
    # Headers and totals are excluded from indexing.
    #
    # ```ballerina
    # type Employee record {| string name; int age; |};
    # Employee first = check table.getRow(0);
    # ```
    #
    # + index - Row index (0-based within data range)
    # + options - Read options
    # + t - Target type descriptor
    # + return - Single row or error
    public isolated function getRow(int index, RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t|Error;

    # Write rows to the table.
    #
    # If the data exceeds the current table size, the table automatically expands.
    # Existing data is overwritten starting from the first data row.
    #
    # ```ballerina
    # Employee[] employees = [{name: "John", age: 30}, {name: "Jane", age: 25}];
    # check table.putRows(employees);
    # ```
    #
    # + data - Data to write (records or arrays)
    # + options - Write options
    # + return - Error if write fails
    public isolated function putRows(Row[] data, *RowWriteOptions options) returns Error?;

    # Check if the table has a total row.
    #
    # + return - true if a total row exists, or an `Error` if the handle is invalid
    public isolated function hasTotalRow() returns boolean|Error;

    # Get the total row values.
    #
    # Returns a map keyed by column name. Each value binds to its natural cell value —
    # a whole number to `int`, a fractional number to `decimal`, a date/time to an ISO
    # string — or `()` for a blank total cell.
    #
    # ```ballerina
    # if check table.hasTotalRow() {
    #     map<xlsx:CellValue?> totals = check table.getTotalRow();
    #     io:println("Total salary: ", totals["Salary"]);
    # }
    # ```
    #
    # + t - Result map type descriptor; leave at the default `map<CellValue?>` (the intended
    #       target). A narrower target (e.g. `map<int>`) only succeeds if every total cell fits it.
    # + return - Map of column names to total values, or error if no total row
    public isolated function getTotalRow(typedesc<map<CellValue?>> t = <>) returns t|Error;

    # Rename the table.
    #
    # The new name must be unique within the workbook.
    #
    # + newName - New table name
    # + return - Error if rename fails (e.g., name already exists)
    public isolated function rename(string newName) returns Error?;

    # Resize the table to a new range.
    #
    # The new range must include at least one header row and one data row.
    # This is for manual resizing; use putRows() for automatic expansion.
    #
    # + newRange - New table range as a `CellRange` record or an A1-notation string
    # + return - Error if resize fails (e.g., invalid range, overlap)
    public isolated function resize(CellRange|string newRange) returns Error?;
};

# Concrete implementation of `Table`. Not exported — instances are vended
# from `Workbook` and `Sheet` methods (`getTable`, `createTable`, etc.).
isolated class TableImpl {
    *Table;

    public isolated function getName() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getDisplayName() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getSheetName() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getRange() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getCellRange() returns CellRange|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getDataRange() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getDataCellRange() returns CellRange|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getRowCount() returns int|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getColumnCount() returns int|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getHeaders() returns string[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getRows(RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getRow(int index, RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function putRows(Row[] data, *RowWriteOptions options) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function hasTotalRow() returns boolean|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getTotalRow(typedesc<map<CellValue?>> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function rename(string newName) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function resize(CellRange|string newRange) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;
}
