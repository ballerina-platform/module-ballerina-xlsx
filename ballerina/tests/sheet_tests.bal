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
import ballerina/time;

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
    if result is Error {
        test:assertTrue(result.message().includes("not found in sheet"),
                "Error message should mention the header was not found in sheet");
    }
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetColumnNilableWithBlank() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    // Header at row 0; data rows 1-3, with row 2's cell intentionally left blank.
    check sheet.setCell(0, 0, "name");
    check sheet.setCell(1, 0, "Alice");
    check sheet.setCell(3, 0, "Charlie");
    string?[] names = check sheet.getColumn("name");
    test:assertEquals(names.length(), 3, "Should read 3 data rows");
    test:assertEquals(names[0], "Alice");
    test:assertTrue(names[1] is (), "Blank cell should read as nil");
    test:assertEquals(names[2], "Charlie");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetColumnBroadCellValue() returns error? {
    // A CellValue?[] (broad) target must preserve natural types, not collapse to strings.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    CellValue?[] ints = check sheet.getColumn("intCol");
    test:assertEquals(ints.length(), 1, "natural_types has one data row");
    test:assertEquals(ints, [42], "Numeric column under CellValue? bound → int");
    CellValue?[] decimals = check sheet.getColumn("decimalCol");
    test:assertEquals(decimals.length(), 1, "natural_types has one data row");
    test:assertEquals(decimals, [3.14d], "Fractional column under CellValue? bound → decimal");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetColumnStringFallback() returns error? {
    // The mirror of the broad case: a pinned string[] target must COLLAPSE genuinely
    // typed cells to their string forms, not return ints/decimals/dates.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    string[] ints = check sheet.getColumn("intCol");
    test:assertEquals(ints, ["42"], "Whole-number cell under string[] → \"42\"");
    string[] decimals = check sheet.getColumn("decimalCol");
    test:assertEquals(decimals, ["3.14"], "Fractional cell under string[] → \"3.14\"");
    string[] dates = check sheet.getColumn("dateCol");
    test:assertEquals(dates, ["2026-05-28"], "Date cell under string[] → ISO string");
    check wb.close();
}

// =============================================================================
// getCell
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetGetCellNaturalTypes() returns error? {
    // Genuinely typed cells must bind to their natural Ballerina types via getCell.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Row 1 holds the typed data cells (row 0 is the header row).
    test:assertEquals(check sheet.getCell(1, 0), 42, "Whole number → int");
    test:assertEquals(check sheet.getCell(1, 1), 3.14d, "Fractional → decimal");
    test:assertEquals(check sheet.getCell(1, 2), true, "Boolean → boolean");
    test:assertEquals(check sheet.getCell(1, 3), "2026-05-28", "Date → ISO string");
    check wb.close();
}

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
function testSheetGetCellDateTime() returns error? {
    // A datetime cell must bind to its ISO string form with the time component preserved.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    test:assertEquals(check sheet.getCell(1, 4), "2026-05-28 14:30:00", "Datetime → ISO string with time");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellBlankIsNil() returns error? {
    // Build a sheet with a known blank cell at (5, 5)
    Workbook wb = new;
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
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, "Header");
    check sheet.setCell(1, 2, 42);
    anydata v00 = check sheet.getCell(0, 0);
    anydata v12 = check sheet.getCell(1, 2);
    test:assertEquals(v00, "Header");
    // A whole-number cell binds to its natural type: int.
    test:assertEquals(v12, 42);
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetCellTypedValues() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, 3.14d);
    check sheet.setCell(0, 1, true);
    time:Date d = {year: 2026, month: 5, day: 28};
    check sheet.setCell(0, 2, d);
    test:assertEquals(check sheet.getCell(0, 0), 3.14d, "Decimal → decimal");
    test:assertEquals(check sheet.getCell(0, 1), true, "Boolean → boolean");
    test:assertEquals(check sheet.getCell(0, 2), "2026-05-28", "Date → ISO string");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetCellByAddress() returns error? {
    Workbook wb = new;
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
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    Error? result = sheet.setCellByAddress("not_a1_address", "value");
    test:assertTrue(result is Error, "Invalid A1 address should return Error");
    if result is Error {
        test:assertTrue(result.message().includes("Invalid cell address"),
                "Error message should mention the invalid cell address");
    }
    check wb.close();
}

// =============================================================================
// setRow
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetSetRowWithStringArray() returns error? {
    Workbook wb = new;
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
    Workbook wb = new;
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
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    // No header row; setRow with a record should error out
    SheetTestEmployee unmatched = {Name: "Alice", Age: 30, Department: "Eng"};
    Error? result = sheet.setRow(0, unmatched);
    test:assertTrue(result is Error, "setRow with record needs an existing header row");
    if result is Error {
        test:assertTrue(result.message().includes("requires an existing header row"),
                "Error message should mention that a header row is required");
    }
    check wb.close();
}

// =============================================================================
// setColumn
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetSetColumnByName() returns error? {
    Workbook wb = new;
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
    // The header row must remain intact and the values must land under the "Bonus" column.
    test:assertEquals(result[0], ["Name", "Bonus"], "Header row should be unchanged");
    string[] bonus = check sheet.getColumn("Bonus");
    test:assertEquals(bonus, ["1000", "2000", "1500"], "Bonus column should hold the written values");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetColumnByIndex() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"], ["3", "4"]]);
    // Replace column 1 starting at the data row
    check sheet.setColumn(1, [99, 100]);
    string[][] result = check sheet.getRows();
    test:assertEquals(result[1][1], "99");
    test:assertEquals(result[2][1], "100");
    // The header row and the untouched column 0 must remain intact.
    test:assertEquals(result[0], ["A", "B"], "Header row should be unchanged");
    test:assertEquals(result[1][0], "1", "Column 0 should be untouched");
    test:assertEquals(result[2][0], "3", "Column 0 should be untouched");
    check wb.close();
}

// =============================================================================
// deleteRow
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetDeleteRowShiftsSubsequentUp() returns error? {
    Workbook wb = new;
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
    Workbook wb = new;
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
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A"], ["B"]]);
    Error? result = sheet.deleteRow(99);
    test:assertTrue(result is Error, "Out-of-range delete should return Error");
    if result is Error {
        test:assertTrue(result.message().includes("out of range for deletion"),
                "Error message should mention the row is out of range for deletion");
    }
    check wb.close();
}

// =============================================================================
// rename
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetRename() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("OldName");
    check sheet.rename("NewName");
    test:assertEquals(check sheet.getName(), "NewName");
    test:assertTrue(check wb.hasSheet("NewName"));
    test:assertFalse(check wb.hasSheet("OldName"));
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetRenameToDuplicate() returns error? {
    Workbook wb = new;
    _ = check wb.createSheet("First");
    Sheet second = check wb.createSheet("Second");
    Error? result = second.rename("First");
    test:assertTrue(result is Error, "Rename to existing sheet name should error");
    if result is Error {
        test:assertTrue(result.message().includes("already exists"),
                "Error message should mention the sheet name already exists");
    }
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
    Workbook wb = new;
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
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    map<CellValue?>[] rows = [
        {"Name": "Alice", "Age": 30},
        {"Name": "Bob", "Age": 25}
    ];
    check sheet.putRows(rows);
    string[][] reread = check sheet.getRows();
    test:assertEquals(reread.length(), 3, "Should have header + 2 data rows");
    string[] headerRow = reread[0];
    test:assertEquals(headerRow.length(), 2, "Header row should hold both keys");
    test:assertTrue(headerRow.indexOf("Name") != (), "Header row should contain Name");
    test:assertTrue(headerRow.indexOf("Age") != (), "Header row should contain Age");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSetRowMapInlineLiteral() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["Name", "Age"], ["Alice", "1"]]);
    map<CellValue?> replacement = {"Name": "Charlie", "Age": 99};
    check sheet.setRow(1, replacement);
    string[][] reread = check sheet.getRows();
    test:assertEquals(reread[0], ["Name", "Age"], "Header row should remain intact after setRow");
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

// =============================================================================
// getRow bounds
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSheetGetRowOutOfRange() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // employees.xlsx has 3 data rows (indices 0-2 relative to the data start row).
    string[]|Error tooHigh = sheet.getRow(99);
    test:assertTrue(tooHigh is Error, "Row index beyond the data range must return Error");
    if tooHigh is Error {
        test:assertTrue(tooHigh.message().includes("out of range"),
                "Error should mention the row is out of range");
    }

    string[]|Error negative = sheet.getRow(-1);
    test:assertTrue(negative is Error, "Negative row index must return Error");

    check wb.close();
}
