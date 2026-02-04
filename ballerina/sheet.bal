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
    # + index - Row index (0-based, relative to data start row)
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
}
