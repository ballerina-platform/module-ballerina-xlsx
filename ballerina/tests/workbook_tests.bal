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

import ballerina/file;
import ballerina/io;
import ballerina/test;

// =============================================================================
// WORKBOOK CREATION TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testOpenExistingWorkbook() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");

    test:assertTrue(wb.getSheetCount() > 0, "Should have at least one sheet");
    test:assertEquals(wb.getSheetCount(), 3, "multi_sheet.xlsx has 3 sheets");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateNewWorkbook() returns error? {
    Workbook wb = check new;

    test:assertEquals(wb.getSheetCount(), 0, "New workbook should have no sheets");
    check wb.close();
}

// =============================================================================
// SHEET ACCESS TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testGetSheetNames() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");

    string[] names = wb.getSheetNames();
    test:assertEquals(names.length(), 3, "Should have 3 sheet names");
    test:assertEquals(names[0], "Sheet1", "First sheet should be 'Sheet1'");
    test:assertEquals(names[1], "Sheet2", "Second sheet should be 'Sheet2'");
    test:assertEquals(names[2], "Sheet3", "Third sheet should be 'Sheet3'");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetByName() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");

    Sheet sheet = check wb.getSheet("Sheet2");
    test:assertEquals(sheet.getName(), "Sheet2", "Sheet name should match");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetByIndex() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");

    // Get first sheet (index 0)
    Sheet sheet0 = check wb.getSheet(0);
    test:assertEquals(sheet0.getName(), "Sheet1", "First sheet should be 'Sheet1'");

    // Get second sheet (index 1)
    Sheet sheet1 = check wb.getSheet(1);
    test:assertEquals(sheet1.getName(), "Sheet2", "Second sheet should be 'Sheet2'");

    // Get third sheet (index 2)
    Sheet sheet2 = check wb.getSheet(2);
    test:assertEquals(sheet2.getName(), "Sheet3", "Third sheet should be 'Sheet3'");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateSheet() returns error? {
    Workbook wb = check new;

    test:assertEquals(wb.getSheetCount(), 0, "Initially no sheets");

    Sheet sheet = check wb.createSheet("TestSheet");
    test:assertEquals(sheet.getName(), "TestSheet", "Sheet name should match");
    test:assertEquals(wb.getSheetCount(), 1, "Should have one sheet");

    // Create another sheet
    Sheet sheet2 = check wb.createSheet("AnotherSheet");
    test:assertEquals(sheet2.getName(), "AnotherSheet", "Second sheet name");
    test:assertEquals(wb.getSheetCount(), 2, "Should have two sheets");

    check wb.close();
}

// =============================================================================
// SHEET NOT FOUND TESTS (NEGATIVE)
// =============================================================================

@test:Config {
    groups: ["workbook", "negative"]
}
function testGetSheetNotFoundByName() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheet("NonExistentSheet");
    test:assertTrue(result is SheetNotFoundError, "Should return SheetNotFoundError");
    if result is SheetNotFoundError {
        test:assertTrue(result.message().includes("not found"),
            "Error message should mention not found");
    }
    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testGetSheetNotFoundByIndex() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheet(99);
    test:assertTrue(result is SheetNotFoundError, "Should return SheetNotFoundError");
    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testGetSheetNegativeIndex() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheet(-1);
    test:assertTrue(result is SheetNotFoundError, "Should return SheetNotFoundError for negative index");
    check wb.close();
}

// =============================================================================
// WORKBOOK SAVE TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testWorkbookSave() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");

    string[][] data = [["Name", "Value"], ["Test", "123"]];
    check sheet.putRows(data);

    string tempFile = getTempFilePath("workbook_save");
    check wb.saveAs(tempFile);
    check wb.close();

    // Verify file was created and has content
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by reading back
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 2, "Should have 2 rows");
    test:assertEquals(parsed[0][0], "Name", "First header");
    test:assertEquals(parsed[1][0], "Test", "First data");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook"]
}
function testWorkbookSaveMultipleSheets() returns error? {
    Workbook wb = check new;

    // Create first sheet
    Sheet sheet1 = check wb.createSheet("Sales");
    string[][] salesInput = [["Product", "Amount"], ["Widget", "100"]];
    check sheet1.putRows(salesInput);

    // Create second sheet
    Sheet sheet2 = check wb.createSheet("Inventory");
    string[][] inventoryInput = [["Item", "Count"], ["Gadget", "50"]];
    check sheet2.putRows(inventoryInput);

    string tempFile = getTempFilePath("workbook_multi_save");
    check wb.saveAs(tempFile);
    check wb.close();

    // Verify by opening and checking both sheets
    Workbook wb2 = check new(tempFile);
    test:assertEquals(wb2.getSheetCount(), 2, "Should have 2 sheets");

    Sheet sales = check wb2.getSheet("Sales");
    string[][] salesData = check sales.getRows();
    test:assertEquals(salesData[0][0], "Product", "Sales header");

    Sheet inventory = check wb2.getSheet("Inventory");
    string[][] inventoryData = check inventory.getRows();
    test:assertEquals(inventoryData[0][0], "Item", "Inventory header");

    check wb2.close();
    check removeTempFile(tempFile);
}

// =============================================================================
// SHEET DATA OPERATIONS TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRows() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    string[][] rows = check sheet.getRows();
    assertStringArrayEquals(rows, EXPECTED_SIMPLE_DATA, "Sheet getRows");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRowsAsRecords() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);

    Employee[] employees = check sheet.getRows();
    assertEmployeesEqual(employees, EXPECTED_EMPLOYEES);
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRowByIndex() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // Get single row as record
    Employee employee = check sheet.getRow(0);  // First data row (after header)
    test:assertEquals(employee.name, "John Doe", "First employee name");
    test:assertEquals(employee.age, 30, "First employee age");
    test:assertEquals(employee.department, "Engineering", "First employee department");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRowByIndexAsString() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // Get single row as string array
    string[] row = check sheet.getRow(0);  // First data row
    test:assertEquals(row[0], "John", "First cell");
    test:assertEquals(row[1], "30", "Second cell");
    test:assertEquals(row[2], "New York", "Third cell");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetPutRows() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("TestSheet");

    string[][] data = [
        ["Col1", "Col2", "Col3"],
        ["A", "B", "C"],
        ["D", "E", "F"]
    ];

    check sheet.putRows(data);

    // Read back
    RowReadOptions opts = {headerRowIndex: 0, dataStartRowIndex: 0};
    string[][] result = check sheet.getRows(opts);
    test:assertEquals(result.length(), 3, "Should have 3 rows");
    assertStringArrayEquals(result, data, "PutRows then getRows");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetPutRowsAsRecords() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Employees");

    Employee[] employees = [
        {name: "Alice", age: 28, department: "Engineering"},
        {name: "Bob", age: 32, department: "Marketing"}
    ];

    check sheet.putRows(employees);

    // Read back
    Employee[] result = check sheet.getRows();
    test:assertEquals(result.length(), 2, "Should have 2 employees");
    test:assertEquals(result[0].name, "Alice", "First employee name");
    test:assertEquals(result[1].name, "Bob", "Second employee name");

    check wb.close();
}

// =============================================================================
// SHEET METADATA TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetName() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");

    Sheet sheet = check wb.getSheet("Sheet2");
    test:assertEquals(sheet.getName(), "Sheet2", "Sheet name should be 'Sheet2'");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetUsedRange() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    string usedRange = sheet.getUsedRange();
    test:assertTrue(usedRange.length() > 0, "Should have used range");
    // simple.xlsx has 4 rows x 3 columns, so range should be like "A1:C4"
    test:assertTrue(usedRange.startsWith("A1"), "Range should start at A1");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetUsedCellRange() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    CellRange? range = sheet.getUsedCellRange();
    test:assertTrue(range != (), "Should have used cell range");

    if range != () {
        // simple.xlsx has 4 rows x 3 columns starting at A1
        test:assertEquals(range.firstRowIndex, 0, "First row should be 0");
        test:assertEquals(range.firstColumnIndex, 0, "First column should be 0");
        test:assertEquals(range.lastRowIndex, 3, "Last row should be 3 (4 rows: 0-3)");
        test:assertEquals(range.lastColumnIndex, 2, "Last column should be 2 (3 columns: 0-2)");
    }

    check wb.close();
}

@test:Config {
    groups: ["workbook", "edge"]
}
function testEmptySheetGetUsedCellRange() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    Sheet sheet = check wb.getSheet(0);

    CellRange? range = sheet.getUsedCellRange();
    test:assertEquals(range, (), "Empty sheet should return nil");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetRowCount() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    int rowCount = sheet.getRowCount();
    test:assertEquals(rowCount, 4, "simple.xlsx has 4 rows (header + 3 data)");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetColumnCount() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);

    int colCount = sheet.getColumnCount();
    test:assertEquals(colCount, 3, "simple.xlsx has 3 columns");

    check wb.close();
}

// =============================================================================
// WORKBOOK LIFECYCLE TESTS
// =============================================================================

@test:Config {
    groups: ["workbook", "lifecycle"]
}
function testWorkbookOpenModifySave() returns error? {
    // First create a file
    string tempFile = getTempFilePath("lifecycle");
    string[][] initialData = [["Original", "Data"]];
    check writeSheet(initialData, tempFile);

    // Open, modify, save
    Workbook wb = check new(tempFile);
    Sheet sheet = check wb.getSheet(0);

    // Add more data
    string[][] newData = [["New", "Row"]];
    check sheet.putRows(newData, startRowIndex = 1);  // Append below existing

    check wb.saveAs(tempFile);
    check wb.close();

    // Verify modifications
    string[][] result = check parseSheet(tempFile);
    test:assertTrue(result.length() >= 2, "Should have original + new data");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook", "lifecycle"]
}
function testWorkbookMultipleOperations() returns error? {
    Workbook wb = check new;

    // Create multiple sheets
    Sheet sheet1 = check wb.createSheet("Sheet1");
    Sheet sheet2 = check wb.createSheet("Sheet2");
    Sheet sheet3 = check wb.createSheet("Sheet3");

    // Write data to each (explicit string[][] typing required)
    string[][] input1 = [["A", "B"], ["1", "2"]];
    string[][] input2 = [["C", "D"], ["3", "4"]];
    string[][] input3 = [["E", "F"], ["5", "6"]];
    check sheet1.putRows(input1);
    check sheet2.putRows(input2);
    check sheet3.putRows(input3);

    // Verify sheet count
    test:assertEquals(wb.getSheetCount(), 3, "Should have 3 sheets");

    // Save and close
    string tempFile = getTempFilePath("multi_ops");
    check wb.saveAs(tempFile);
    check wb.close();

    // Reopen and verify
    Workbook wb2 = check new(tempFile);
    test:assertEquals(wb2.getSheetCount(), 3, "Should still have 3 sheets");

    Sheet s1 = check wb2.getSheet("Sheet1");
    string[][] data1 = check s1.getRows();
    test:assertEquals(data1[0][0], "A", "Sheet1 data preserved");

    Sheet s3 = check wb2.getSheet("Sheet3");
    string[][] data3 = check s3.getRows();
    test:assertEquals(data3[0][0], "E", "Sheet3 data preserved");

    check wb2.close();
    check removeTempFile(tempFile);
}

// =============================================================================
// SHEET GETROWS WITH OPTIONS TESTS
// =============================================================================

@test:Config {
    groups: ["workbook", "options"]
}
function testSheetGetRowsWithHeaderOption() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "complex_headers.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // complex_headers.xlsx: row 0=title, row 1=metadata, row 2=headers, row 3+=data
    RowReadOptions opts = {
        headerRowIndex: 2,
        dataStartRowIndex: 3
    };
    string[][] rows = check sheet.getRows(opts);

    // When dataStartRowIndex is specified, only data rows are returned
    test:assertEquals(rows.length(), 2, "Should have 2 data rows");
    test:assertEquals(rows[0][0], "Item1", "First data row");
    test:assertEquals(rows[1][0], "Item2", "Second data row");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "options"]
}
function testSheetGetRowsWithFormulaModeText() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "formulas.xlsx");
    Sheet sheet = check wb.getSheet(0);

    RowReadOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check sheet.getRows(opts);

    // In TEXT mode, at least one cell should have a formula string
    boolean hasFormula = false;
    foreach string[] row in rows {
        foreach string cell in row {
            if cell.startsWith("=") {
                hasFormula = true;
                break;
            }
        }
    }
    test:assertTrue(hasFormula, "Should have formula text");

    check wb.close();
}

// =============================================================================
// EDGE CASE TESTS
// =============================================================================

@test:Config {
    groups: ["workbook", "edge"]
}
function testEmptySheetMetadata() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // Empty sheet should have 0 rows/columns
    int rowCount = sheet.getRowCount();
    int colCount = sheet.getColumnCount();

    test:assertEquals(rowCount, 0, "Empty sheet should have 0 rows");
    test:assertEquals(colCount, 0, "Empty sheet should have 0 columns");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "edge"]
}
function testEmptySheetGetRows() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    Sheet sheet = check wb.getSheet(0);

    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 0, "Empty sheet should return empty array");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "edge"]
}
function testSingleCellSheet() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "edge_single_cell.xlsx");
    Sheet sheet = check wb.getSheet(0);

    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 1, "Should have 1 row");
    test:assertEquals(rows[0].length(), 1, "Should have 1 column");
    test:assertEquals(rows[0][0], "alone", "Cell value");

    check wb.close();
}

// =============================================================================
// DELETE SHEET TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testDeleteSheetByName() returns error? {
    Workbook wb = check new;

    // Create multiple sheets
    _ = check wb.createSheet("Sheet1");
    _ = check wb.createSheet("Sheet2");
    _ = check wb.createSheet("Sheet3");
    test:assertEquals(wb.getSheetCount(), 3, "Should have 3 sheets");

    // Delete middle sheet
    check wb.deleteSheet("Sheet2");
    test:assertEquals(wb.getSheetCount(), 2, "Should have 2 sheets after delete");

    // Verify remaining sheets
    string[] names = wb.getSheetNames();
    test:assertEquals(names[0], "Sheet1", "First sheet should be Sheet1");
    test:assertEquals(names[1], "Sheet3", "Second sheet should be Sheet3");

    // Verify Sheet2 is not found
    Sheet|SheetNotFoundError result = wb.getSheet("Sheet2");
    test:assertTrue(result is SheetNotFoundError, "Sheet2 should not be found");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testDeleteSheetByIndex() returns error? {
    Workbook wb = check new;

    // Create multiple sheets
    _ = check wb.createSheet("First");
    _ = check wb.createSheet("Second");
    _ = check wb.createSheet("Third");
    test:assertEquals(wb.getSheetCount(), 3, "Should have 3 sheets");

    // Delete first sheet (index 0)
    check wb.deleteSheet(0);
    test:assertEquals(wb.getSheetCount(), 2, "Should have 2 sheets after delete");

    // Verify remaining sheets
    string[] names = wb.getSheetNames();
    test:assertEquals(names[0], "Second", "First sheet should now be Second");
    test:assertEquals(names[1], "Third", "Second sheet should now be Third");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testDeleteSheetByNameNotFound() returns error? {
    Workbook wb = check new;
    _ = check wb.createSheet("Sheet1");

    Error? result = wb.deleteSheet("NonExistent");
    test:assertTrue(result is SheetNotFoundError,
            "Not-found case must surface as the specific SheetNotFoundError subtype");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testDeleteSheetByIndexOutOfRange() returns error? {
    Workbook wb = check new;
    _ = check wb.createSheet("Sheet1");

    Error? result = wb.deleteSheet(99);
    test:assertTrue(result is SheetNotFoundError,
            "Out-of-range index must surface as the specific SheetNotFoundError subtype");

    check wb.close();
}

// =============================================================================
// SAVE AND SAVEAS TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testSaveInMemoryWorkbookError() returns error? {
    // In-memory workbooks have no source path, so save() should fail
    Workbook wb = check new;
    _ = check wb.createSheet("Data");

    Error? result = wb.save();
    test:assertTrue(result is Error, "save() should fail for in-memory workbook");
    if result is Error {
        test:assertTrue(result.message().includes("no source file"),
            "Error should mention no source file");
    }

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSaveAsUpdateSourcePath() returns error? {
    // saveAs() should update the source path, so subsequent save() works
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    string[][] data = [["Name", "Value"], ["Test", "123"]];
    check sheet.putRows(data);

    string tempFile = getTempFilePath("saveas_test");

    // First saveAs
    check wb.saveAs(tempFile);

    // Modify data
    string[][] moreData = [["More", "Data"]];
    check sheet.putRows(moreData, startRowIndex = 2);

    // Now save() should work (to the same path)
    check wb.save();
    check wb.close();

    // Verify by reading back
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have 3 rows");
    test:assertEquals(parsed[2][0], "More", "Added row should exist");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook"]
}
function testNewWorkbookAndSaveAs() returns error? {
    // saveAs() writes the file and sets the source path, so a later save() works too
    string tempFile = getTempFilePath("saveas_test");

    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    string[][] data = [["Header1", "Header2"], ["A", "B"]];
    check sheet.putRows(data);

    check wb.saveAs(tempFile);
    check wb.close();

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify content
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 2, "Should have 2 rows");
    test:assertEquals(parsed[0][0], "Header1", "First header");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook"]
}
function testOpenFileAndSaveOverwrites() returns error? {
    // First create a file
    string tempFile = getTempFilePath("openfile_overwrite");
    string[][] initialData = [["Original", "Data"]];
    check writeSheet(initialData, tempFile);

    // Open and modify
    Workbook wb = check new(tempFile);
    Sheet sheet = check wb.getSheet(0);
    string[][] newData = [["Modified", "Content"]];
    check sheet.putRows(newData);

    // save() should overwrite the original file
    check wb.save();
    check wb.close();

    // Verify modifications persisted
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed[0][0], "Modified", "Should be overwritten");

    check removeTempFile(tempFile);
}

// =============================================================================
// hasSheet / toBytes TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testWorkbookHasSheetExisting() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");
    test:assertTrue(wb.hasSheet("Sheet1"), "Sheet1 should exist");
    test:assertTrue(wb.hasSheet("Sheet2"), "Sheet2 should exist");
    test:assertTrue(wb.hasSheet("Sheet3"), "Sheet3 should exist");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testWorkbookHasSheetMissing() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "multi_sheet.xlsx");
    test:assertFalse(wb.hasSheet("DoesNotExist"), "Non-existent sheet should return false");
    test:assertFalse(wb.hasSheet(""), "Empty name should return false");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testWorkbookToBytesRoundTrip() returns error? {
    // Build a workbook in memory and serialize to bytes
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    string[][] data = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]];
    check sheet.putRows(data);
    byte[] bytes = check wb.toBytes();
    check wb.close();

    test:assertTrue(bytes.length() > 0, "Serialized bytes should be non-empty");

    // Re-open from the bytes and verify the data
    Workbook wb2 = check new(bytes);
    test:assertEquals(wb2.getSheetCount(), 1, "Should have 1 sheet");
    test:assertTrue(wb2.hasSheet("Data"), "Sheet 'Data' should be present in re-opened workbook");
    Sheet sheet2 = check wb2.getSheet("Data");
    string[][] parsed = check sheet2.getRows();
    test:assertEquals(parsed, data, "Round-tripped data should match original");
    check wb2.close();
}

// =============================================================================
// Handle invalidation tests
// =============================================================================
// Vended Sheet/Table handles are invalidated on Workbook.close() / deleteSheet /
// deleteTable. Method calls on invalidated handles return a typed Error instead
// of touching a closed POI workbook.

@test:Config {
    groups: ["workbook"]
}
function testUseSheetAfterCloseReturnsError() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);
    check wb.close();
    string[][]|Error result = sheet.getRows();
    test:assertTrue(result is Error, "Sheet.getRows after Workbook.close should return Error");
}

@test:Config {
    groups: ["workbook"]
}
function testUseSheetAfterDeleteReturnsError() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Doomed");
    _ = check wb.createSheet("Survivor");  // can't delete last sheet
    check wb.deleteSheet("Doomed");
    string[][]|Error result = sheet.getRows();
    test:assertTrue(result is Error, "Sheet.getRows after deleteSheet should return Error");
    check wb.close();
}

// =============================================================================
// Sheet name validation + case-insensitive lookup + refuse-delete-last
// =============================================================================

@test:Config {groups: ["workbook"]}
function testCreateSheetInvalidName() returns error? {
    Workbook wb = check new;
    Sheet|Error result = wb.createSheet("Has/Slash");
    test:assertTrue(result is Error, "Forbidden char '/' must be rejected");
    check wb.close();
}

@test:Config {groups: ["workbook"]}
function testCreateSheetExceeds31Chars() returns error? {
    Workbook wb = check new;
    Sheet|Error result = wb.createSheet("ThisSheetNameIsWayTooLongForExcel32");
    test:assertTrue(result is Error, "Names over 31 chars must be rejected");
    check wb.close();
}

@test:Config {groups: ["workbook"]}
function testCreateSheetEmptyName() returns error? {
    Workbook wb = check new;
    Sheet|Error result = wb.createSheet("");
    test:assertTrue(result is Error, "Empty sheet name must be rejected");
    check wb.close();
}

@test:Config {groups: ["workbook"]}
function testCaseInsensitiveSheetLookup() returns error? {
    Workbook wb = check new;
    _ = check wb.createSheet("Sales");
    // Excel sheet names are case-insensitive on lookup.
    Sheet sheet = check wb.getSheet("sales");
    test:assertEquals(sheet.getName(), "Sales");
    test:assertTrue(wb.hasSheet("SALES"), "hasSheet should be case-insensitive");
    check wb.close();
}

@test:Config {groups: ["workbook"]}
function testDeleteLastSheetRefused() returns error? {
    Workbook wb = check new;
    _ = check wb.createSheet("OnlyOne");
    Error? result = wb.deleteSheet("OnlyOne");
    test:assertTrue(result is Error, "Deleting the last sheet must be refused");
    check wb.close();
}

@test:Config {groups: ["workbook"]}
function testGetNameAfterCloseReturnsError() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheet(0);
    check wb.close();

    string|error nameResult = trap sheet.getName();
    test:assertTrue(nameResult is error,
            "Sheet.getName after Workbook.close should error, not silently succeed");
}

// =============================================================================
// Non-XLSX content surfaces as ParseError, not FileNotFoundError
// =============================================================================
// WorkbookHandle.openWorkbookFromPath used to map every IOException, including
// parse failures from WorkbookFactory.create, to FileNotFoundError. The fix
// discriminates: missing file → FileNotFoundError; unreadable content → ParseError.

@test:Config {groups: ["workbook"]}
function testNonXlsxFileReturnsParseError() returns error? {
    string tempFile = getTempFilePath("not_xlsx");
    check io:fileWriteString(tempFile, "this is not a valid xlsx file");

    Workbook|Error result = new(tempFile);
    test:assertTrue(result is ParseError,
            "Non-XLSX file content must surface as ParseError");
    test:assertFalse(result is FileNotFoundError,
            "Non-XLSX file must not be misclassified as FileNotFoundError");

    check removeTempFile(tempFile);
}
