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
# xlsx:Workbook wb = check new("sales.xlsx");
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
public type Table object {

    # Get the name of the table.
    #
    # Table names are unique across the entire workbook.
    #
    # + return - Table name
    public function getName() returns string;

    # Get the display name of the table.
    #
    # The display name is what appears in Excel's UI. It may differ from the internal name.
    #
    # + return - Table display name
    public function getDisplayName() returns string;

    # Get the name of the sheet containing this table.
    #
    # + return - Sheet name
    public function getSheetName() returns string;

    # Get the full range of the table.
    #
    # Returns the entire table range including headers and totals row (if present).
    # All indices are 0-based.
    #
    # + return - CellRange representing the full table area
    public function getRange() returns CellRange;

    # Get the data range of the table.
    #
    # Returns only the data rows, excluding headers and totals row.
    # All indices are 0-based.
    #
    # + return - CellRange representing only the data area
    public function getDataRange() returns CellRange;

    # Get the number of data rows in the table.
    #
    # Returns only data rows, excluding header and totals row.
    #
    # + return - Data row count
    public function getRowCount() returns int;

    # Get the number of columns in the table.
    #
    # + return - Column count
    public function getColumnCount() returns int;

    # Get the column header names.
    #
    # Returns an array of header names in column order.
    #
    # + return - Array of header strings
    public function getHeaders() returns string[];

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
    # + t - Target type descriptor
    # + return - Array of data rows or error
    public function getRows(RowReadOptions options = {}, typedesc<Data> t = <>)
            returns t|Error;

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
    public function getRow(int index, RowReadOptions options = {}, typedesc<Row> t = <>)
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
    public function putRows(Data data, *RowWriteOptions options) returns Error?;

    # Check if the table has a totals row.
    #
    # + return - true if totals row exists
    public function hasTotalsRow() returns boolean;

    # Get the totals row values.
    #
    # Returns a map where keys are column names and values are the totals.
    #
    # ```ballerina
    # if table.hasTotalsRow() {
    #     map<anydata> totals = check table.getTotalsRow();
    #     io:println("Total salary: ", totals["Salary"]);
    # }
    # ```
    #
    # + return - Map of column names to totals values, or error if no totals row
    public function getTotalsRow() returns map<anydata>|Error;

    # Rename the table.
    #
    # The new name must be unique within the workbook.
    #
    # + newName - New table name
    # + return - Error if rename fails (e.g., name already exists)
    public function rename(string newName) returns Error?;

    # Resize the table to a new range.
    #
    # The new range must include at least one header row and one data row.
    # This is for manual resizing; use putRows() for automatic expansion.
    #
    # + newRange - New table range
    # + return - Error if resize fails (e.g., invalid range, overlap)
    public function resize(CellRange newRange) returns Error?;
};

# Concrete implementation of `Table`. Not exported — instances are vended
# from `Workbook` and `Sheet` methods (`getTable`, `createTable`, etc.).
class TableImpl {
    *Table;

    public function getName() returns string = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getDisplayName() returns string = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getSheetName() returns string = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getRange() returns CellRange = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getDataRange() returns CellRange = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getRowCount() returns int = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getColumnCount() returns int = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getHeaders() returns string[] = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getRows(RowReadOptions options = {}, typedesc<Data> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getRow(int index, RowReadOptions options = {}, typedesc<Row> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function putRows(Data data, *RowWriteOptions options) returns Error? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function hasTotalsRow() returns boolean = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function getTotalsRow() returns map<anydata>|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function rename(string newName) returns Error? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;

    public function resize(CellRange newRange) returns Error? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.TableHandle"
    } external;
}