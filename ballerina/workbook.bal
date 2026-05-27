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
# ```ballerina
# // Empty in-memory workbook
# xlsx:Workbook empty = check new;
#
# // Open an existing file
# xlsx:Workbook fromFile = check new("report.xlsx");
# string[] sheets = fromFile.getSheetNames();
# xlsx:Sheet sheet = check fromFile.getSheet("Sales");
# // ... modify data ...
# check fromFile.save();   // Overwrites the original file
# check fromFile.close();
#
# // Open from a byte array (e.g., HTTP payload)
# xlsx:Workbook fromBytes = check new(sourceBytes);
# ```
public class Workbook {

    # Initialize a workbook.
    #
    # - `()` (default) — create an empty in-memory workbook with no file association.
    # - `string` — open an existing XLSX file at the given path.
    # - `byte[]` — open an XLSX workbook from an in-memory byte array.
    #
    # ```ballerina
    # xlsx:Workbook empty = check new;
    # xlsx:Workbook fromFile = check new("data.xlsx");
    # xlsx:Workbook fromBytes = check new(sourceBytes);
    # ```
    #
    # + input - File path, byte array, or `()` for an empty workbook
    # + return - Error if the input cannot be opened
    public function init((string|byte[])? input = ()) returns Error? {
        if input is string {
            check self.initFromPath(input);
        } else if input is byte[] {
            check self.initFromBytes(input);
        } else {
            check self.initNew();
        }
    }

    # Initialize workbook from a file path.
    # + path - Path to the XLSX file
    # + return - Error if file not found or parsing fails
    function initFromPath(string path) returns Error? = @java:Method {
        name: "openWorkbookFromPath",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Initialize workbook from a byte array.
    # + sourceBytes - XLSX content as a byte array
    # + return - Error if the byte array is not a valid XLSX workbook
    function initFromBytes(byte[] sourceBytes) returns Error? = @java:Method {
        name: "openWorkbookFromBytes",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Initialize a new empty workbook.
    #
    # + return - Error if creation fails
    function initNew() returns Error? = @java:Method {
        name: "createNewWorkbook",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Set the source path for the workbook.
    # + path - Path to associate with this workbook
    # + return - Error if setting path fails
    function setSourcePathNative(string path) returns Error? = @java:Method {
        name: "setSourcePath",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get all sheet names in the workbook.
    #
    # + return - Array of sheet names
    public function getSheetNames() returns string[] = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get the number of sheets in the workbook.
    #
    # + return - Sheet count
    public function getSheetCount() returns int = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function hasSheet(string name) returns boolean = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function getSheet(string|int target) returns Sheet|SheetNotFoundError {
        SheetImpl sheet = new;
        if target is string {
            check self.getSheetByNameNative(sheet, target);
        } else {
            check self.getSheetByIndexNative(sheet, target);
        }
        return sheet;
    }

    function getSheetByNameNative(Sheet sheet, string name) returns SheetNotFoundError? = @java:Method {
        name: "getSheet",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    function getSheetByIndexNative(Sheet sheet, int index) returns SheetNotFoundError? = @java:Method {
        name: "getSheetByIndex",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function createSheet(string name) returns Sheet|Error {
        SheetImpl sheet = new;
        check self.createSheetNative(sheet, name);
        return sheet;
    }

    function createSheetNative(Sheet sheet, string name) returns Error? = @java:Method {
        name: "createSheet",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function deleteSheet(string|int target) returns Error? {
        if target is string {
            return self.deleteSheetByNameNative(target);
        }
        return self.deleteSheetByIndexNative(target);
    }

    function deleteSheetByNameNative(string name) returns Error? = @java:Method {
        name: "deleteSheet",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    function deleteSheetByIndexNative(int index) returns Error? = @java:Method {
        name: "deleteSheetByIndex",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook to its source file.
    #
    # Overwrites the file the workbook was opened from (`new(path)`) or the
    # path most recently passed to `saveAs(path)`. Returns an error for
    # in-memory workbooks (`new`) that haven't yet been saved to a path.
    #
    # ```ballerina
    # xlsx:Workbook wb = check new("data.xlsx");
    # // ... modify ...
    # check wb.save();  // Overwrites data.xlsx
    # ```
    #
    # + return - Error if no source path or save fails
    public function save() returns Error? = @java:Method {
        name: "saveToSource",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function saveAs(string path) returns Error? {
        check self.saveToPathNative(path);
        check self.setSourcePathNative(path);
    }

    function saveToPathNative(string path) returns Error? = @java:Method {
        name: "saveToPath",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function toBytes() returns byte[]|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Close the workbook and release resources.
    #
    # Always call this when done with the workbook to free memory.
    #
    # + return - Error if close fails
    public function close() returns Error? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function getTable(string name) returns Table|TableNotFoundError = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
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
    public function getAllTables() returns Table[]|Error = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;
}
