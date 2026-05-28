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

# Represents an Excel workbook.
#
# A workbook contains one or more sheets and provides methods to
# access sheets, create new sheets, delete sheets, and save to files.
#
# Construct an empty in-memory workbook directly with `new`. Use the
# module-level `xlsx:fromFile` and `xlsx:fromBytes` factory functions
# to open an existing XLSX from a path or byte array.
#
# ```ballerina
# // Empty in-memory workbook
# xlsx:Workbook empty = check new;
#
# // Open an existing file
# xlsx:Workbook wb = check xlsx:fromFile("report.xlsx");
# string[] sheets = wb.getSheetNames();
# xlsx:Sheet sheet = check wb.getSheet("Sales");
# // ... modify data ...
# check wb.save();   // Overwrites the original file
# check wb.close();
#
# // Open from a byte array (e.g., HTTP payload)
# xlsx:Workbook fromBytes = check xlsx:fromBytes(sourceBytes);
# ```
public isolated class Workbook {

    # Initialize an empty in-memory workbook.
    #
    # No file association; `save()` errors until a path is set via `saveAs(path)`.
    # To open an existing workbook, use `xlsx:fromFile(path)` or `xlsx:fromBytes(bytes)`.
    #
    # + return - Error if workbook creation fails
    public isolated function init() returns Error? {
        check self.initNew();
    }

    # Initialize a new empty workbook.
    #
    # + return - Error if creation fails
    isolated function initNew() returns Error? = @java:Method {
        name: "createNewWorkbook",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get all sheet names in the workbook.
    #
    # + return - Array of sheet names
    public isolated function getSheetNames() returns string[] = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get the number of sheets in the workbook.
    #
    # + return - Sheet count
    public isolated function getSheetCount() returns int = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Check whether a sheet with the given name exists in the workbook.
    #
    # ```ballerina
    # if wb.hasSheet("Sales") {
    #     xlsx:Sheet s = check wb.getSheet("Sales");
    # }
    # ```
    #
    # + name - Sheet name
    # + return - `true` if the sheet exists, `false` otherwise
    public isolated function hasSheet(string name) returns boolean = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get a sheet by name or index.
    #
    # ```ballerina
    # xlsx:Sheet byName = check workbook.getSheet("Sales");
    # xlsx:Sheet byIndex = check workbook.getSheet(0);
    # ```
    #
    # + target - Sheet name (string) or 0-based index (int)
    # + return - Sheet instance or SheetNotFoundError if not found
    public isolated function getSheet(string|int target) returns Sheet|SheetNotFoundError {
        return target is string ? self.getSheetByName(target) : self.getSheetByIndex(target);
    }

    isolated function getSheetByName(string name) returns Sheet|SheetNotFoundError = @java:Method {
        name: "getSheet",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    isolated function getSheetByIndex(int index) returns Sheet|SheetNotFoundError = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Create a new sheet in the workbook.
    #
    # ```ballerina
    # xlsx:Sheet newSheet = check workbook.createSheet("Report");
    # check newSheet.putRows(data);
    # ```
    #
    # + name - Name for the new sheet
    # + return - New sheet instance or Error if name already exists
    public isolated function createSheet(string name) returns Sheet|Error = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Delete a sheet by name or index.
    #
    # ```ballerina
    # check workbook.deleteSheet("TempData");
    # check workbook.deleteSheet(0);            // Delete first sheet
    # ```
    #
    # + target - Sheet name (string) or 0-based index (int)
    # + return - `SheetNotFoundError` if the sheet doesn't exist, or another `Error`
    #           (e.g., refusing to delete the last sheet — Excel requires at least one)
    public isolated function deleteSheet(string|int target) returns Error? {
        if target is string {
            return self.deleteSheetByNameNative(target);
        }
        return self.deleteSheetByIndexNative(target);
    }

    isolated function deleteSheetByNameNative(string name) returns Error? = @java:Method {
        name: "deleteSheet",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    isolated function deleteSheetByIndexNative(int index) returns Error? = @java:Method {
        name: "deleteSheetByIndex",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook to its source file.
    #
    # Overwrites the file the workbook was opened from (`xlsx:fromFile(path)`)
    # or the path most recently passed to `saveAs(path)`. Returns an error for
    # in-memory workbooks (`new`) that haven't yet been saved to a path.
    #
    # ```ballerina
    # xlsx:Workbook wb = check xlsx:fromFile("data.xlsx");
    # // ... modify ...
    # check wb.save();  // Overwrites data.xlsx
    # ```
    #
    # + return - Error if no source path or save fails
    public isolated function save() returns Error? = @java:Method {
        name: "saveToSource",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook to a new location.
    #
    # After calling saveAs(), subsequent calls to save() will
    # write to this new location.
    #
    # ```ballerina
    # xlsx:Workbook wb = check new;
    # xlsx:Sheet sheet = check wb.createSheet("Data");
    # check sheet.putRows(data);
    # check wb.saveAs("output.xlsx");
    # // Now wb.save() would write to output.xlsx
    # ```
    #
    # + path - Path to save the XLSX file
    # + return - Error if save fails
    public isolated function saveAs(string path) returns Error? {
        check self.saveToPathNative(path);
    }

    isolated function saveToPathNative(string path) returns Error? = @java:Method {
        name: "saveToPath",
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Serialize the workbook to a byte array.
    #
    # Useful for returning the workbook as an HTTP response body, embedding it
    # in a larger payload, or any flow that needs the bytes without writing
    # to a file.
    #
    # ```ballerina
    # byte[] bytes = check wb.toBytes();
    # // ... use the bytes ...
    # ```
    #
    # + return - XLSX bytes or Error if serialization fails
    public isolated function toBytes() returns byte[]|Error = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Close the workbook and release resources.
    #
    # Always call this when done with the workbook to free memory.
    #
    # + return - Error if close fails
    public isolated function close() returns Error? = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    // =============================================================================
    // TABLE METHODS
    // =============================================================================

    # Get a table by name from anywhere in the workbook.
    #
    # Table names are unique across the entire workbook, so no sheet
    # specification is needed.
    #
    # ```ballerina
    # xlsx:Table empTable = check wb.getTable("EmployeeTable");
    # Employee[] employees = check empTable.getRows();
    # ```
    #
    # + name - Table name
    # + return - Table or error if not found
    public isolated function getTable(string name) returns Table|TableNotFoundError = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get all tables across all sheets in the workbook.
    #
    # ```ballerina
    # xlsx:Table[] allTables = check wb.getAllTables();
    # foreach xlsx:Table t in allTables {
    #     io:println("Table: ", t.getName(), " in sheet: ", t.getSheetName());
    # }
    # ```
    #
    # + return - Array of all tables (may be empty), or Error on retrieval failure
    public isolated function getAllTables() returns Table[]|Error = @java:Method {
        'class: "io.ballerina.stdlib.xlsx.xlsx.WorkbookHandle"
    } external;
}
