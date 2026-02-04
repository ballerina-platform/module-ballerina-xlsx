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
# access sheets, create new sheets, and save to files.
#
# ```ballerina
# // Open existing workbook
# xlsx:Workbook wb = check new("report.xlsx");
#
# // Create new empty workbook
# xlsx:Workbook wb = check new();
#
# // Work with sheets
# string[] sheets = wb.getSheetNames();
# xlsx:Sheet sheet = check wb.getSheet("Sales");
#
# // Save and close
# check wb.save("updated.xlsx");
# check wb.close();
# ```
public class Workbook {

    # Initialize a workbook.
    #
    # + path - Path to XLSX file, or nil to create new empty workbook
    # + return - Error if file not found or initialization fails
    public function init(string? path = ()) returns Error? {
        if path is string {
            check self.initFromPath(path);
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

    # Initialize a new empty workbook.
    #
    # + return - Error if creation fails
    function initNew() returns Error? = @java:Method {
        name: "createNewWorkbook",
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

    # Save the workbook to a file.
    #
    # ```ballerina
    # check workbook.save("output.xlsx");
    # ```
    #
    # + path - Path to save the XLSX file
    # + return - Error if save fails
    public function save(string path) returns Error? = @java:Method {
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
}
