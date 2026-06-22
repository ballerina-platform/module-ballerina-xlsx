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

# A worksheet in a workbook, with methods to read and write rows, columns, cells, and tables.
# Obtained from a `Workbook` (for example `getSheet` or `createSheet`); not constructed directly.
public type Sheet isolated object {

    # Get the name of the sheet.
    #
    # + return - The sheet name, or an error
    public isolated function getName() returns string|Error;

    # Get the used range of the sheet in A1 notation, such as "A1:D50".
    #
    # + return - The used range, or an error
    public isolated function getUsedRange() returns string|Error;

    # Get the used cell range as a structured record, or nil if the sheet is empty.
    #
    # + return - The used range as 0-based indices, nil if the sheet is empty, or an error
    public isolated function getUsedCellRange() returns CellRange?|Error;

    # Get the number of rows with data.
    #
    # + return - The row count, or an error
    public isolated function getRowCount() returns int|Error;

    # Get the number of columns with data.
    #
    # + return - The column count, or an error
    public isolated function getColumnCount() returns int|Error;

    # Read all rows from the sheet as records, maps, or a string grid.
    #
    # + options - Read options
    # + t - Target row type
    # + return - The rows, or an error
    public isolated function getRows(ParseOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error;

    # Read a single row by index as a record, map, or string array.
    #
    # + index - 0-based index within the data window, so `getRow(0)` is the first data row
    # + options - Read options
    # + t - Target row type
    # + return - The row, or an error
    public isolated function getRow(int index, RowParseOptions options = {}, typedesc<Row> t = <>)
            returns t|Error;

    # Write rows to the sheet (records, maps, or string arrays).
    #
    # By default rows are appended below the existing data; `sheetWriteMode` selects another
    # disposition.
    #
    # + data - Rows to write
    # + options - Write options
    # + return - An error if the write fails
    public isolated function putRows(Row[] data, *WriteOptions options) returns Error?;

    # Get a column of values by header name or 0-based index.
    #
    # + columnRef - Column header name, or 0-based index
    # + options - Read options
    # + t - Target cell type (`CellValue`, which includes `()` for blank cells)
    # + return - The column values, or an error
    public isolated function getColumn(string|int columnRef, ColumnParseOptions options = {},
            typedesc<CellValue> t = <>) returns t[]|Error;

    # Read a single cell, bound to the target type.
    #
    # The target type drives the binding: the default `CellValue` yields the cell's natural value,
    # while a `time:Civil` / `time:Date` / `time:TimeOfDay` or scalar target yields that type.
    #
    # + rowIndex - 0-based row index
    # + columnIndex - 0-based column index
    # + t - Target cell type (default: `CellValue`)
    # + return - The cell value, `()` for a blank cell when the target allows it, or an error
    public isolated function getCell(int rowIndex, int columnIndex, typedesc<CellValue> t = <>)
            returns t|Error;

    # Write a single row at the given 0-based row index.
    #
    # By default the row is overwritten; `sheetWriteMode` selects another disposition. For a record
    # or map, values align to columns by header name, using the header at `options.headerRowIndex`.
    #
    # + rowIndex - 0-based row index
    # + data - Row data (`string[]`, record, or `map<CellValue>`)
    # + options - Single-row write options
    # + return - An error if the write fails
    public isolated function setRow(int rowIndex, Row data, *RowWriteOptions options)
            returns Error?;

    # Write a column of values by header name or 0-based index.
    #
    # Values are written into successive rows below the header row.
    #
    # + columnRef - Column header name, or 0-based index
    # + data - Column values
    # + return - An error if the write fails
    public isolated function setColumn(string|int columnRef, CellValue[] data)
            returns Error?;

    # Write a single cell by 0-based row and column index.
    #
    # + rowIndex - 0-based row index
    # + columnIndex - 0-based column index
    # + value - Cell value
    # + return - An error if the write fails
    public isolated function setCell(int rowIndex, int columnIndex, CellValue value)
            returns Error?;

    # Write a single cell by A1-notation address.
    #
    # + cellAddress - Cell address in A1 notation, such as "A1" or "B12"
    # + value - Cell value
    # + return - An error if the address is invalid or the write fails
    public isolated function setCellByAddress(string cellAddress, CellValue value)
            returns Error?;

    # Delete a row from the sheet; subsequent rows shift up by one.
    #
    # + index - 0-based row index to delete
    # + return - An error if the delete fails
    public isolated function deleteRow(int index) returns Error?;

    # Rename the sheet.
    #
    # The new name must follow Excel rules (at most 31 characters, none of `\ / ? * [ ] :`) and be
    # unique in the workbook.
    #
    # + newName - New sheet name
    # + return - An error if the name is invalid or already taken
    public isolated function rename(string newName) returns Error?;

    # Get a table on this sheet by name.
    #
    # + name - Table name
    # + return - The table, or a `TableNotFoundError` if it does not exist
    public isolated function getTable(string name) returns Table|Error;

    # Get all tables on this sheet.
    #
    # + return - The tables (may be empty), or an error on failure
    public isolated function getTables() returns Table[]|Error;

    # Create a table over an existing range.
    #
    # The range must include a header row. If `headers` is not given, the first row is used.
    #
    # + name - Unique table name (across the workbook)
    # + range - Table range as a `CellRange` or A1-notation string
    # + headers - Optional custom headers; if omitted, the first row is used
    # + return - The created table, or an error
    public isolated function createTable(string name, CellRange|string range, string[]? headers = ())
            returns Table|Error;

    # Write data and wrap it in a new table, computing the range automatically.
    #
    # The table always has a header row: field names (or `@xlsx:Name`) for records, keys for maps,
    # or the first row for `string[][]`.
    #
    # + name - Unique table name (across the workbook)
    # + data - Data to write (records, maps, or string arrays)
    # + startRowIndex - Starting row for the table (default: 0)
    # + startColumnIndex - Starting column for the table (default: 0)
    # + return - The created table, or an error
    public isolated function createTableFromData(string name, Row[] data,
            int startRowIndex = 0, int startColumnIndex = 0)
            returns Table|Error;

    # Delete a table from this sheet. The underlying data is preserved.
    #
    # + name - Table name to delete
    # + return - A `TableNotFoundError` if the table does not exist, or another error
    public isolated function deleteTable(string name) returns Error?;
};

# Concrete implementation of `Sheet`. Not exported; instances are vended
# from `Workbook` methods such as `getSheet` and `createSheet`.
isolated class SheetImpl {
    *Sheet;

    public isolated function getName() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getUsedRange() returns string|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getUsedCellRange() returns CellRange?|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getRowCount() returns int|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getColumnCount() returns int|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getRows(ParseOptions options = {}, typedesc<Row> t = <>)
            returns t[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getRow(int index, RowParseOptions options = {}, typedesc<Row> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function putRows(Row[] data, *WriteOptions options) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getColumn(string|int columnRef, ColumnParseOptions options = {},
            typedesc<CellValue> t = <>) returns t[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getCell(int rowIndex, int columnIndex, typedesc<CellValue> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function setRow(int rowIndex, Row data, *RowWriteOptions options)
            returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function setColumn(string|int columnRef, CellValue[] data)
            returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function setCell(int rowIndex, int columnIndex, CellValue value)
            returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function setCellByAddress(string cellAddress, CellValue value)
            returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function deleteRow(int index) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function rename(string newName) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getTable(string name) returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function getTables() returns Table[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function createTable(string name, CellRange|string range, string[]? headers = ())
            returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function createTableFromData(string name, Row[] data,
            int startRowIndex = 0, int startColumnIndex = 0)
            returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;

    public isolated function deleteTable(string name) returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.SheetHandle"
    } external;
}
