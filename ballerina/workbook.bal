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

// =============================================================================
// WORKBOOK FACTORY FUNCTIONS
// =============================================================================

# Opens an existing XLSX file.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:openFile("report.xlsx");
# string[] sheets = wb.getSheetNames();
# check wb.close();
# ```
#
# + path - Path to the existing XLSX file
# + return - Workbook handle or FileNotFoundError if file doesn't exist
public function openFile(string path) returns Workbook|Error {
    Workbook wb = new;
    check wb.initFromPath(path);
    return wb;
}

# Creates a new XLSX file at the specified path.
#
# The file is not written until save() or saveAs() is called.
# Calling save() will write to the specified path.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:createFile("report.xlsx");
# xlsx:Sheet sheet = check wb.createSheet("Data");
# check sheet.putRows(data);
# check wb.save();  // Writes to report.xlsx
# check wb.close();
# ```
#
# + path - Path where the file will be created when saved
# + return - Workbook handle
public function createFile(string path) returns Workbook|Error {
    Workbook wb = new;
    check wb.initNew();
    check wb.setSourcePathNative(path);
    return wb;
}

# Creates a new in-memory workbook with no file association.
#
# Use saveAs(path) to write the workbook to a file. Calling save()
# without a prior saveAs() will return an error.
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:createWorkbook();
# xlsx:Sheet sheet = check wb.createSheet("Data");
# check sheet.putRows(data);
# check wb.saveAs("output.xlsx");  // Must use saveAs for in-memory workbooks
# check wb.close();
# ```
#
# + return - Workbook handle
public function createWorkbook() returns Workbook|Error {
    Workbook wb = new;
    check wb.initNew();
    return wb;
}

// =============================================================================
// WORKBOOK CLASS
// =============================================================================

# Represents an Excel workbook.
#
# A workbook contains one or more sheets and provides methods to
# access sheets, create new sheets, delete sheets, and save to files.
#
# Use the factory functions to create workbook instances:
# - `xlsx:openFile(path)` - Open existing file
# - `xlsx:createFile(path)` - Create file (writes on save)
# - xlsx:createWorkbook() - Create in-memory workbook
#
# ```ballerina
# xlsx:Workbook wb = check xlsx:openFile("report.xlsx");
# string[] sheets = wb.getSheetNames();
# xlsx:Sheet sheet = check wb.getSheet("Sales");
# // ... modify data ...
# check wb.save();  // Overwrites original file
# check wb.close();
# ```
public class Workbook {

    # Initialize workbook from a file path.
    # + path - Path to the XLSX file
    # + return - Error if file not found or parsing fails
    function initFromPath(string path) returns Error? = @java:Method {
        name: "openWorkbookFromPath",
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

    # Get a sheet by name.
    #
    # ```ballerina
    # xlsx:Sheet sheet = check workbook.getSheet("Sales");
    # ```
    #
    # + name - Sheet name
    # + return - Sheet instance or SheetNotFoundError
    public function getSheet(string name) returns Sheet|SheetNotFoundError {
        Sheet sheet = new;
        check self.getSheetNative(sheet, name);
        return sheet;
    }

    function getSheetNative(Sheet sheet, string name) returns SheetNotFoundError? = @java:Method {
        name: "getSheet",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Get a sheet by index.
    #
    # ```ballerina
    # xlsx:Sheet firstSheet = check workbook.getSheetByIndex(0);
    # ```
    #
    # + index - Sheet index (0-based)
    # + return - Sheet instance or SheetNotFoundError if index out of range
    public function getSheetByIndex(int index) returns Sheet|SheetNotFoundError {
        Sheet sheet = new;
        check self.getSheetByIndexNative(sheet, index);
        return sheet;
    }

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
        Sheet sheet = new;
        check self.createSheetNative(sheet, name);
        return sheet;
    }

    function createSheetNative(Sheet sheet, string name) returns Error? = @java:Method {
        name: "createSheet",
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Delete a sheet by name.
    #
    # ```ballerina
    # check workbook.deleteSheet("TempData");
    # ```
    #
    # + name - Name of the sheet to delete
    # + return - SheetNotFoundError if sheet doesn't exist
    public function deleteSheet(string name) returns SheetNotFoundError? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Delete a sheet by index.
    #
    # ```ballerina
    # check workbook.deleteSheetByIndex(0);  // Delete first sheet
    # ```
    #
    # + index - Index of the sheet to delete (0-based)
    # + return - SheetNotFoundError if index out of range
    public function deleteSheetByIndex(int index) returns SheetNotFoundError? = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;

    # Save the workbook to its source file.
    #
    # Overwrites the original file for workbooks created with `openFile()`
    # or `createFile()`. Returns error for in-memory workbooks created
    # with createWorkbook() - use saveAs() instead.
    #
    # ```ballerina
    # xlsx:Workbook wb = check xlsx:openFile("data.xlsx");
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
    # xlsx:Workbook wb = check xlsx:createWorkbook();
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
    # xlsx:Table[] allTables = wb.getAllTables();
    # foreach xlsx:Table t in allTables {
    #     io:println("Table: ", t.getName(), " in sheet: ", t.getSheetName());
    # }
    # ```
    #
    # + return - Array of all tables (may be empty)
    public function getAllTables() returns Table[] = @java:Method {
        'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
    } external;
}
