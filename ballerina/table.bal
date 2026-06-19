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

# An Excel Table (ListObject) in a worksheet, with automatic header handling, an optional totals
# row, and auto-resize on write. Table names are unique across the workbook.
#
# Obtained from a `Workbook` or `Sheet` (for example `getTable` or `createTable`); not constructed
# directly.
#
# ```ballerina
# xlsx:Table empTable = check wb.getTable("EmployeeTable");
# Employee[] employees = check empTable.getRows();
# check empTable.putRows(newEmployees);
# ```
public type Table isolated object {

    # Get the name of the table. Table names are unique across the workbook.
    #
    # + return - The table name, or an error
    public isolated function getName() returns string|Error;

    # Get the display name of the table, as shown in the Excel UI.
    #
    # + return - The display name, or an error
    public isolated function getDisplayName() returns string|Error;

    # Get the name of the sheet that holds this table.
    #
    # + return - The sheet name, or an error
    public isolated function getSheetName() returns string|Error;

    # Get the full table range, including the header and totals row, in A1 notation.
    #
    # + return - The range in A1 notation, such as "A1:D10", or an error
    public isolated function getRange() returns string|Error;

    # Get the full table range, including the header and totals row, as a `CellRange` (0-based).
    #
    # + return - The full table range, or an error
    public isolated function getCellRange() returns CellRange|Error;

    # Get the data range, excluding the header and totals row, in A1 notation.
    #
    # + return - The data range in A1 notation, or an error
    public isolated function getDataRange() returns string|Error;

    # Get the data range, excluding the header and totals row, as a `CellRange` (0-based).
    #
    # + return - The data range, or an error
    public isolated function getDataCellRange() returns CellRange|Error;

    # Get the number of data rows, excluding the header and totals row.
    #
    # + return - The data row count, or an error
    public isolated function getRowCount() returns int|Error;

    # Get the number of columns in the table.
    #
    # + return - The column count, or an error
    public isolated function getColumnCount() returns int|Error;

    # Get the column header names, in column order.
    #
    # + return - The header names, or an error
    public isolated function getHeaders() returns string[]|Error;

    # Read all data rows from the table as records, maps, or a string grid.
    #
    # The header and any totals row are excluded.
    #
    # + options - Table read options
    # + t - Target row type
    # + return - The data rows, or an error
    public isolated function getRows(TableParseOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error;

    # Read a single data row by index as a record, map, or string array.
    #
    # + index - 0-based index within the data range, so `getRow(0)` is the first data row
    # + options - Table read options
    # + t - Target row type
    # + return - The row, or an error
    public isolated function getRow(int index, TableRowParseOptions options = {}, typedesc<Row> t = <>)
            returns t|Error;

    # Write rows to the table, resizing its data range to fit.
    #
    # By default the data is replaced; `tableWriteMode = APPEND` adds rows below it instead.
    # A resize that would overlap another table fails with a `TableOverlapError`.
    #
    # + data - Rows to write (records, maps, or string arrays)
    # + options - Table write options
    # + return - An error if the write fails
    public isolated function putRows(Row[] data, *TableWriteOptions options) returns Error?;

    # Check whether the table has a totals row.
    #
    # + return - Whether a totals row exists, or an error
    public isolated function hasTotalRow() returns boolean|Error;

    # Get the totals row as a map keyed by column name.
    #
    # Each value binds to its natural cell value, or `()` for a blank total cell.
    #
    # + t - Result map type (default: `map<CellValue>`); a narrower type binds only if every cell fits
    # + return - The totals by column name, or an error if there is no totals row
    public isolated function getTotalRow(typedesc<map<CellValue>> t = <>) returns t|Error;

    # Rename the table. The new name must be unique in the workbook.
    #
    # + newName - New table name
    # + return - An error if the rename fails, for example if the name is taken
    public isolated function rename(string newName) returns Error?;

    # Resize the table to a new range, which must include a header row and a data row.
    # For automatic resizing on write, use `putRows` instead.
    #
    # + newRange - New table range as a `CellRange` or an A1-notation string
    # + return - An error if the range is invalid or overlaps another table
    public isolated function resize(CellRange|string newRange) returns Error?;

    # Delete a data row by 0-based index; the table shrinks and rows below move up.
    #
    # A table must keep at least one data row, so the last one cannot be deleted.
    #
    # + index - 0-based index within the data range, so 0 is the first data row
    # + return - An error if the index is out of range, it is the last data row, or the shrink
    #            would disrupt another table
    public isolated function deleteRow(int index) returns Error?;
};

# Concrete implementation of `Table`. Not exported; instances are vended
# from `Workbook` and `Sheet` methods such as `getTable` and `createTable`.
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

    public isolated function getRows(TableParseOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getRow(int index, TableRowParseOptions options = {}, typedesc<Row> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function putRows(Row[] data, *TableWriteOptions options) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function hasTotalRow() returns boolean|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function getTotalRow(typedesc<map<CellValue>> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function rename(string newName) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function resize(CellRange|string newRange) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;

    public isolated function deleteRow(int index) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.TableHandle"
    } external;
}
