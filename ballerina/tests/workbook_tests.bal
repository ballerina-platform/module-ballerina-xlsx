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
import ballerina/test;

// =============================================================================
// WORKBOOK CREATION TESTS
// =============================================================================

@test:Config {
    groups: ["workbook"]
}
function testOpenExistingWorkbook() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    test:assertTrue(wb.getSheetCount() > 0, "Should have at least one sheet");
    test:assertEquals(wb.getSheetCount(), 3, "multi_sheet.xlsx has 3 sheets");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateNewWorkbook() returns error? {
    Workbook wb = check new Workbook();

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    Sheet sheet = check wb.getSheet("Sheet2");
    test:assertEquals(sheet.getName(), "Sheet2", "Sheet name should match");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetByIndex() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    // Get first sheet (index 0)
    Sheet sheet0 = check wb.getSheetByIndex(0);
    test:assertEquals(sheet0.getName(), "Sheet1", "First sheet should be 'Sheet1'");

    // Get second sheet (index 1)
    Sheet sheet1 = check wb.getSheetByIndex(1);
    test:assertEquals(sheet1.getName(), "Sheet2", "Second sheet should be 'Sheet2'");

    // Get third sheet (index 2)
    Sheet sheet2 = check wb.getSheetByIndex(2);
    test:assertEquals(sheet2.getName(), "Sheet3", "Third sheet should be 'Sheet3'");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateSheet() returns error? {
    Workbook wb = check new Workbook();

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheetByIndex(99);
    test:assertTrue(result is SheetNotFoundError, "Should return SheetNotFoundError");
    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testGetSheetNegativeIndex() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheetByIndex(-1);
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
    Workbook wb = check new Workbook();
    Sheet sheet = check wb.createSheet("Data");

    string[][] data = [["Name", "Value"], ["Test", "123"]];
    check sheet.putRows(data);

    string tempFile = getTempFilePath("workbook_save");
    check wb.save(tempFile);
    check wb.close();

    // Verify file was created and has content
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by reading back
    string[][] parsed = check parse(tempFile);
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
    Workbook wb = check new Workbook();

    // Create first sheet
    Sheet sheet1 = check wb.createSheet("Sales");
    string[][] salesInput = [["Product", "Amount"], ["Widget", "100"]];
    check sheet1.putRows(salesInput);

    // Create second sheet
    Sheet sheet2 = check wb.createSheet("Inventory");
    string[][] inventoryInput = [["Item", "Count"], ["Gadget", "50"]];
    check sheet2.putRows(inventoryInput);

    string tempFile = getTempFilePath("workbook_multi_save");
    check wb.save(tempFile);
    check wb.close();

    // Verify by opening and checking both sheets
    Workbook wb2 = check new Workbook(tempFile);
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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string[][] rows = check sheet.getRows();
    assertStringArrayEquals(rows, EXPECTED_SIMPLE_DATA, "Sheet getRows");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRowsAsRecords() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    Employee[] employees = check sheet.getRows();
    assertEmployeesEqual(employees, EXPECTED_EMPLOYEES);
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRowByIndex() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

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
    Workbook wb = check new Workbook();
    Sheet sheet = check wb.createSheet("TestSheet");

    string[][] data = [
        ["Col1", "Col2", "Col3"],
        ["A", "B", "C"],
        ["D", "E", "F"]
    ];

    check sheet.putRows(data);

    // Read back
    RowReadOptions opts = {headerRow: 0, dataStartRow: 0};
    string[][] result = check sheet.getRows(opts);
    test:assertEquals(result.length(), 3, "Should have 3 rows");
    assertStringArrayEquals(result, data, "PutRows then getRows");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetPutRowsAsRecords() returns error? {
    Workbook wb = check new Workbook();
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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    Sheet sheet = check wb.getSheet("Sheet2");
    test:assertEquals(sheet.getName(), "Sheet2", "Sheet name should be 'Sheet2'");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetUsedRange() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string usedRange = sheet.getUsedRange();
    test:assertTrue(usedRange.length() > 0, "Should have used range");
    // simple.xlsx has 4 rows x 3 columns, so range should be like "A1:C4"
    test:assertTrue(usedRange.startsWith("A1"), "Range should start at A1");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetRowCount() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    int rowCount = sheet.getRowCount();
    test:assertEquals(rowCount, 4, "simple.xlsx has 4 rows (header + 3 data)");

    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadataGetColumnCount() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

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
    check write(initialData, tempFile);

    // Open, modify, save
    Workbook wb = check new Workbook(tempFile);
    Sheet sheet = check wb.getSheetByIndex(0);

    // Add more data
    string[][] newData = [["New", "Row"]];
    check sheet.putRows(newData, startRow = 1);  // Append below existing

    check wb.save(tempFile);
    check wb.close();

    // Verify modifications
    string[][] result = check parse(tempFile);
    test:assertTrue(result.length() >= 2, "Should have original + new data");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook", "lifecycle"]
}
function testWorkbookMultipleOperations() returns error? {
    Workbook wb = check new Workbook();

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
    check wb.save(tempFile);
    check wb.close();

    // Reopen and verify
    Workbook wb2 = check new Workbook(tempFile);
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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "complex_headers.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    // complex_headers.xlsx: row 0=title, row 1=metadata, row 2=headers, row 3+=data
    RowReadOptions opts = {
        headerRow: 2,
        dataStartRow: 3
    };
    string[][] rows = check sheet.getRows(opts);

    // When dataStartRow is specified, only data rows are returned
    test:assertEquals(rows.length(), 2, "Should have 2 data rows");
    test:assertEquals(rows[0][0], "Item1", "First data row");
    test:assertEquals(rows[1][0], "Item2", "Second data row");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "options"]
}
function testSheetGetRowsWithFormulaModeText() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "formulas.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

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
    Workbook wb = check new Workbook(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 0, "Empty sheet should return empty array");

    check wb.close();
}

@test:Config {
    groups: ["workbook", "edge"]
}
function testSingleCellSheet() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "edge_single_cell.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 1, "Should have 1 row");
    test:assertEquals(rows[0].length(), 1, "Should have 1 column");
    test:assertEquals(rows[0][0], "alone", "Cell value");

    check wb.close();
}
