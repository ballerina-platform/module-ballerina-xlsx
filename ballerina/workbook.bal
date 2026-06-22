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

# An Excel workbook: a set of sheets, with methods to read, create, delete, and save them.
#
# Create an empty workbook with `new`, or open one with `xlsx:fromFile` / `xlsx:fromBytes`.
# A workbook and the sheets and tables obtained from it are not safe for concurrent mutation.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:fromFile("report.xlsx");
# xlsx:Sheet sheet = check wb.getSheet("Sales");
# check wb.save();
# check wb.close();
# ```
public isolated class Workbook {

    # Create an empty in-memory workbook. Persist it with `saveAs(path)`, since `save()` has no
    # source path yet. To open an existing workbook, use `xlsx:fromFile` or `xlsx:fromBytes`.
    public isolated function init() {
        self.initNew();
    }

    # Initialize a new empty workbook.
    isolated function initNew() = @java:Method {
        name: "createNewWorkbook",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get all sheet names in the workbook.
    #
    # + return - Array of sheet names
    public isolated function getSheetNames() returns string[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get the number of sheets in the workbook.
    #
    # + return - Sheet count
    public isolated function getSheetCount() returns int|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Check whether a sheet with the given name exists.
    #
    # + name - Sheet name
    # + return - Whether the sheet exists, or an error if the workbook is closed
    public isolated function hasSheet(string name) returns boolean|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get a sheet by name or 0-based index.
    #
    # + target - Sheet name, or 0-based index
    # + return - The sheet, or a `SheetNotFoundError` if it does not exist
    public isolated function getSheet(string|int target) returns Sheet|Error {
        return target is string ? self.getSheetByName(target) : self.getSheetByIndex(target);
    }

    isolated function getSheetByName(string name) returns Sheet|Error = @java:Method {
        name: "getSheet",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    isolated function getSheetByIndex(int index) returns Sheet|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Create a new sheet in the workbook.
    #
    # + name - Name for the new sheet
    # + return - The new sheet, or a `SheetExistsError` if the name is taken
    public isolated function createSheet(string name) returns Sheet|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Delete a sheet by name or 0-based index.
    #
    # + target - Sheet name, or 0-based index
    # + return - A `SheetNotFoundError` if the sheet is missing, or an error if it is the last sheet
    public isolated function deleteSheet(string|int target) returns Error? {
        if target is string {
            return self.deleteSheetByNameNative(target);
        }
        return self.deleteSheetByIndexNative(target);
    }

    isolated function deleteSheetByNameNative(string name) returns Error? = @java:Method {
        name: "deleteSheet",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    isolated function deleteSheetByIndexNative(int index) returns Error? = @java:Method {
        name: "deleteSheetByIndex",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook, overwriting the file it was opened from or last saved to with `saveAs`.
    #
    # + return - An error if the workbook has no source path (created with `new`), or if the save fails
    public isolated function save() returns Error? = @java:Method {
        name: "saveToSource",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook to a new path, which then becomes the target of later `save()` calls.
    #
    # + path - Path to write the XLSX file to
    # + return - An error if the save fails
    public isolated function saveAs(string path) returns Error? {
        check self.saveToPathNative(path);
    }

    isolated function saveToPathNative(string path) returns Error? = @java:Method {
        name: "saveToPath",
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Serialize the workbook to a byte array, for example to send as an HTTP response.
    #
    # + return - The XLSX bytes, or an error if serialization fails
    public isolated function toBytes() returns byte[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Close the workbook and release its resources. Call this when done to free memory.
    #
    # + return - An error if the close fails
    public isolated function close() returns Error? = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    // =============================================================================
    // TABLE METHODS
    // =============================================================================

    # Get a table by name from anywhere in the workbook. Table names are unique workbook-wide.
    #
    # + name - Table name
    # + return - The table, or a `TableNotFoundError` if it does not exist
    public isolated function getTable(string name) returns Table|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get all tables across every sheet in the workbook.
    #
    # + return - All tables (may be empty), or an error on failure
    public isolated function getAllTables() returns Table[]|Error = @java:Method {
        'class: "io.ballerina.lib.xlsx.xlsx.WorkbookHandle"
    } external;
}
