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

# Represents a worksheet within an Excel workbook.
#
# A sheet contains rows of data and provides methods to read and write data.
public class Sheet {

    # Get the name of the sheet.
    #
    # + return - Sheet name
    public function getName() returns string = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get the used range of the sheet in A1 notation.
    #
    # The used range is the smallest rectangular area that contains all cells with data.
    # This excludes "ghost rows" - rows that have formatting but no actual data.
    #
    # + return - Range string (e.g., "A1:D50")
    public function getUsedRange() returns string = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get the used cell range of the sheet as a structured record.
    #
    # Returns 0-based row and column indices representing the rectangular area
    # containing all cells with actual data. Returns `nil` if the sheet is empty.
    #
    # This is useful for:
    # - Determining sheet bounds before iterating
    # - Planning data writes at specific positions
    # - Understanding the actual data footprint of a sheet
    #
    # ```ballerina
    # xlsx:Workbook wb = check new("data.xlsx");
    # xlsx:Sheet sheet = check wb.getSheet("Sales");
    # xlsx:CellRange? range = sheet.getUsedCellRange();
    # if range != () {
    #     io:println("Data spans rows ", range.firstRowIndex, " to ", range.lastRowIndex);
    # }
    # ```
    #
    # + return - CellRange record with 0-based indices, or nil if sheet is empty
    public function getUsedCellRange() returns CellRange? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get the number of rows with data.
    #
    # + return - Row count
    public function getRowCount() returns int = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get the number of columns with data.
    #
    # + return - Column count
    public function getColumnCount() returns int = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get all rows from the sheet.
    #
    # Supports reading to:
    # - `string[][]` - Raw string array
    # - `record{}[]` - Array of records
    #
    # ```ballerina
    # // As string array
    # string[][] rows = check sheet.getRows();
    #
    # // As records
    # type Employee record {| string name; int age; |};
    # Employee[] employees = check sheet.getRows();
    # ```
    #
    # + options - Read options
    # + t - Target type descriptor
    # + return - Array of rows or error
    public function getRows(RowReadOptions options = {}, typedesc<anydata[]> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get a single row from the sheet by index.
    #
    # Supports reading to:
    # - `string[]` - Raw string array for the row
    # - `record{}` - Single record
    #
    # ```ballerina
    # // Get row as string array
    # string[] row = check sheet.getRow(5);
    #
    # // Get row as record
    # type Employee record {| string name; int age; |};
    # Employee employee = check sheet.getRow(5);
    # ```
    #
    # + index - Row index (0-based, relative to data start row).
#           Example: If dataStartRowIndex=1, getRow(0) returns Excel row 1, getRow(2) returns Excel row 3.
    # + options - Read options
    # + t - Target type descriptor
    # + return - Single row or error
    public function getRow(int index, RowReadOptions options = {}, typedesc<anydata> t = <>)
            returns t|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Write rows to the sheet.
    #
    # Supports writing from:
    # - `string[][]` - Raw string array
    # - `record{}[]` - Array of records (field names become headers)
    #
    # ```ballerina
    # // Write string array
    # string[][] data = [["Name", "Age"], ["John", "30"]];
    # check sheet.putRows(data);
    #
    # // Write records
    # Employee[] employees = [{name: "John", age: 30}];
    # check sheet.putRows(employees);
    # ```
    #
    # + data - Data to write
    # + options - Write options
    # + return - Error if write fails
    public function putRows(anydata[] data, *RowWriteOptions options) returns Error? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    // =============================================================================
    // TABLE METHODS
    // =============================================================================

    # Get a table from this sheet by name.
    #
    # ```ballerina
    # xlsx:Table empTable = check sheet.getTable("EmployeeTable");
    # ```
    #
    # + name - Table name
    # + return - Table or error if not found
    public function getTable(string name) returns Table|TableNotFoundError = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Get all tables in this sheet.
    #
    # ```ballerina
    # xlsx:Table[] tables = sheet.getTables();
    # foreach xlsx:Table t in tables {
    #     io:println("Table: ", t.getName());
    # }
    # ```
    #
    # + return - Array of tables (may be empty)
    public function getTables() returns Table[] = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Create a new table with the specified range.
    #
    # The range must include at least a header row. If headers are not provided,
    # the first row of the range is used as headers.
    #
    # ```ballerina
    # // Create from range string
    # xlsx:Table t1 = check sheet.createTable("SalesTable", "A1:D10");
    #
    # // Create from CellRange with custom headers
    # xlsx:Table t2 = check sheet.createTable("BonusTable", {
    #     firstRowIndex: 0, lastRowIndex: 5,
    #     firstColumnIndex: 0, lastColumnIndex: 3
    # }, ["Name", "Department", "Amount", "Date"]);
    # ```
    #
    # + name - Unique table name (across workbook)
    # + range - Table range as CellRange record or A1 notation string
    # + headers - Optional custom headers; if not provided, first row is used
    # + return - Created table or error
    public function createTable(string name, CellRange|string range, string[]? headers = ())
            returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Create a new table from data, automatically calculating the range.
    #
    # Writes the data first, then creates a table around it.
    #
    # ```ballerina
    # Employee[] employees = [...];
    # xlsx:Table empTable = check sheet.createTableFromData("EmployeeTable", employees);
    # ```
    #
    # + name - Unique table name (across workbook)
    # + data - Data to write (records or arrays)
    # + startRowIndex - Starting row for the table (default: 0)
    # + startColumnIndex - Starting column for the table (default: 0)
    # + return - Created table or error
    public function createTableFromData(string name, Data data,
            int startRowIndex = 0, int startColumnIndex = 0)
            returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;

    # Delete a table from this sheet.
    #
    # The table structure is removed but the underlying data is preserved.
    #
    # ```ballerina
    # check sheet.deleteTable("OldTable");
    # ```
    #
    # + name - Table name to delete
    # + return - Error if table not found
    public function deleteTable(string name) returns TableNotFoundError? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.SheetHandle"
    } external;
}
