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

import ballerina/test;

// =============================================================================
// Helper record types
// =============================================================================

type SheetTestEmployee record {|
    string Name;
    int Age;
    string Department;
|};

// =============================================================================
// getColumn
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetGetColumnByName() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    string[] names = check sheet.getColumn("name");
    test:assertEquals(names.length(), 3, "Should have 3 names");
    test:assertEquals(names[0], "John Doe");
    test:assertEquals(names[2], "Bob Johnson");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetColumnByIndex() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Column 0 is "name"; pull as string[]
    string[] col0 = check sheet.getColumn(0);
    test:assertEquals(col0.length(), 3);
    test:assertEquals(col0[0], "John Doe");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetColumnMissingHeader() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    string[]|Error result = sheet.getColumn("NonExistentHeader");
    test:assertTrue(result is Error, "Missing header should return Error");
    check wb.close();
}

// =============================================================================
// getCell
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetGetCell() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Row 0 is the header row; cell (0,0) = "name"
    anydata header00 = check sheet.getCell(0, 0);
    test:assertEquals(header00, "name");
    // Row 1, col 0 = first employee's name
    anydata cell10 = check sheet.getCell(1, 0);
    test:assertEquals(cell10, "John Doe");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellBlankIsNil() returns error? {
    // Build a sheet with a known blank cell at (5, 5)
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    anydata blank = check sheet.getCell(5, 5);
    test:assertEquals(blank, (), "Blank cell should return ()");
    check wb.close();
}

// =============================================================================
// setCell / setCellByAddress
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetSetCell() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, "Header");
    check sheet.setCell(1, 2, 42);
    anydata v00 = check sheet.getCell(0, 0);
    anydata v12 = check sheet.getCell(1, 2);
    test:assertEquals(v00, "Header");
    // Numeric cells with an anydata target come back as decimal.
    test:assertEquals(v12, 42d);
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetCellByAddress() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCellByAddress("A1", "Header");
    check sheet.setCellByAddress("D5", 42.5d);
    anydata a1 = check sheet.getCell(0, 0);
    anydata d5 = check sheet.getCell(4, 3);
    test:assertEquals(a1, "Header");
    test:assertEquals(d5, 42.5d);
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetCellByAddressInvalid() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    Error? result = sheet.setCellByAddress("not_a1_address", "value");
    test:assertTrue(result is Error, "Invalid A1 address should return Error");
    check wb.close();
}

// =============================================================================
// setRow
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetSetRowWithStringArray() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["Name", "Age"], ["Old", "1"]]);
    // Overwrite row 1
    check sheet.setRow(1, ["Alice", "30"]);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows[1], ["Alice", "30"]);
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetRowWithRecord() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    SheetTestEmployee[] initial = [
        {Name: "Alice", Age: 30, Department: "Eng"},
        {Name: "Bob", Age: 25, Department: "Sales"}
    ];
    check sheet.putRows(initial);
    // Overwrite row 2 (data row 1, 0-based row 2 because header is row 0)
    SheetTestEmployee replacement = {Name: "Charlie", Age: 40, Department: "Marketing"};
    check sheet.setRow(2, replacement);
    SheetTestEmployee[] reread = check sheet.getRows();
    test:assertEquals(reread[1].Name, "Charlie");
    test:assertEquals(reread[1].Department, "Marketing");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetRowWithRecordRequiresHeaderRow() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    // No header row; setRow with a record should error out
    SheetTestEmployee unmatched = {Name: "Alice", Age: 30, Department: "Eng"};
    Error? result = sheet.setRow(0, unmatched);
    test:assertTrue(result is Error, "setRow with record needs an existing header row");
    check wb.close();
}

// =============================================================================
// setColumn
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetSetColumnByName() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([
        ["Name", "Bonus"],
        ["Alice", "0"],
        ["Bob", "0"],
        ["Carol", "0"]
    ]);
    check sheet.setColumn("Bonus", [1000, 2000, 1500]);
    string[][] result = check sheet.getRows();
    test:assertEquals(result[1][1], "1000");
    test:assertEquals(result[2][1], "2000");
    test:assertEquals(result[3][1], "1500");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetColumnByIndex() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"], ["3", "4"]]);
    // Replace column 1 starting at the data row
    check sheet.setColumn(1, [99, 100]);
    string[][] result = check sheet.getRows();
    test:assertEquals(result[1][1], "99");
    test:assertEquals(result[2][1], "100");
    check wb.close();
}

// =============================================================================
// deleteRow
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetDeleteRowShiftsSubsequentUp() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([
        ["A"],
        ["B"],
        ["C"],
        ["D"]
    ]);
    // Delete row 1 ("B"); rows below shift up
    check sheet.deleteRow(1);
    string[][] result = check sheet.getRows();
    test:assertEquals(result.length(), 3, "Should have 3 rows after deletion");
    test:assertEquals(result[0][0], "A");
    test:assertEquals(result[1][0], "C", "Row originally at index 2 should now be at index 1");
    test:assertEquals(result[2][0], "D");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetDeleteLastRow() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A"], ["B"], ["C"]]);
    check sheet.deleteRow(2);
    string[][] result = check sheet.getRows();
    test:assertEquals(result.length(), 2);
    test:assertEquals(result[1][0], "B");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetDeleteRowOutOfRange() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A"], ["B"]]);
    Error? result = sheet.deleteRow(99);
    test:assertTrue(result is Error, "Out-of-range delete should return Error");
    check wb.close();
}

// =============================================================================
// rename
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetRename() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("OldName");
    check sheet.rename("NewName");
    test:assertEquals(sheet.getName(), "NewName");
    test:assertTrue(wb.hasSheet("NewName"));
    test:assertFalse(wb.hasSheet("OldName"));
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetRenameToDuplicate() returns error? {
    Workbook wb = check new;
    _ = check wb.createSheet("First");
    Sheet second = check wb.createSheet("Second");
    Error? result = second.rename("First");
    test:assertTrue(result is Error, "Rename to existing sheet name should error");
    check wb.close();
}

// =============================================================================
// Public Data/Row union dispatch tests
// =============================================================================
// These tests exercise the public API surface with values whose declared type is
// `Data` or `Row` — the abstract union types. They lock in the native dispatch
// path that unwraps BTypeReferenceType and handles UNION_TAG on element types.

@test:Config {groups: ["sheet"]}
function testPutRowsRecordInlineLiteral() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    // Build a record literal in a way that resolves to the closed record type via the
    // declared array type. The inline-literal form forces contextual typing against
    // `Data` = `Row[]`, sending a union-element BArray into Java.
    SheetTestEmployee[] employees = [
        {Name: "Alice", Age: 30, Department: "Eng"},
        {Name: "Bob", Age: 25, Department: "Sales"}
    ];
    check sheet.putRows(employees);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 3, "Should have header + 2 rows");
    test:assertEquals(rows[0], ["Name", "Age", "Department"]);
    test:assertEquals(rows[1][0], "Alice");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsMapInlineLiteral() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    map<anydata>[] rows = [
        {"Name": "Alice", "Age": 30},
        {"Name": "Bob", "Age": 25}
    ];
    check sheet.putRows(rows);
    string[][] reread = check sheet.getRows();
    test:assertEquals(reread.length(), 3, "Should have header + 2 data rows");
    test:assertTrue(reread[0].indexOf("Name") != (), "Should have Name header");
    test:assertTrue(reread[0].indexOf("Age") != (), "Should have Age header");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSetRowMapInlineLiteral() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["Name", "Age"], ["Alice", "1"]]);
    map<anydata> replacement = {"Name": "Charlie", "Age": 99};
    check sheet.setRow(1, replacement);
    string[][] reread = check sheet.getRows();
    test:assertEquals(reread[1][0], "Charlie", "Row 1 Name should be Charlie");
    test:assertEquals(reread[1][1], "99", "Row 1 Age should be 99");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetRowsWithDataTarget() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Explicit `Row[]` target — the inferred element is the `Row` union itself;
    // native dispatch must handle the UNION element tag and fall back to string[][].
    Row[] rows = check sheet.getRows();
    test:assertTrue(rows is string[][], "Row[] target should fall back to string[][]");
    if rows is string[][] {
        test:assertEquals(rows.length(), 4, "Should have header + 3 data rows");
        test:assertEquals(rows[0], ["name", "age", "department"]);
    }
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetRowWithRowTarget() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Explicit `Row` target — declared type is the union itself.
    // Sheet.getRow uses data-row-relative indexing: getRow(0) is the first row of data.
    Row r = check sheet.getRow(0);
    test:assertTrue(r is string[], "Row target should fall back to string[]");
    if r is string[] {
        test:assertEquals(r[0], "John Doe");
    }
    check wb.close();
}
