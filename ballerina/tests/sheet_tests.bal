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
function testSheetGetColumnCaseInsensitive() returns error? {
    // case_headers.xlsx has mixed-case headers (NAME/AGE/Department). With
    // caseInsensitiveHeaders, a lowercase column reference must still resolve.
    Workbook wb = check fromFile(TEST_DATA_DIR + "case_headers.xlsx");
    Sheet sheet = check wb.getSheet(0);
    ColumnParseOptions opts = {caseInsensitiveHeaders: true};
    string[] names = check sheet.getColumn("name", opts);
    test:assertEquals(names.length(), 2, "Should read 2 data rows");
    test:assertEquals(names[0], "John");
    test:assertEquals(names[1], "Jane");
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
    // A CellValue[] (broad) target must preserve natural types, not collapse to strings.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    CellValue[] ints = check sheet.getColumn("intCol");
    test:assertEquals(ints.length(), 1, "natural_types has one data row");
    test:assertEquals(ints, [42], "Numeric column under CellValue bound → int");
    CellValue[] decimals = check sheet.getColumn("decimalCol");
    test:assertEquals(decimals.length(), 1, "natural_types has one data row");
    test:assertEquals(decimals, [3.14d], "Fractional column under CellValue bound → decimal");
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
    CellValue intCell = check sheet.getCell(1, 0);
    test:assertEquals(intCell, 42, "Whole number → int");
    CellValue decimalCell = check sheet.getCell(1, 1);
    test:assertEquals(decimalCell, 3.14d, "Fractional → decimal");
    CellValue boolCell = check sheet.getCell(1, 2);
    test:assertEquals(boolCell, true, "Boolean → boolean");
    CellValue dateCell = check sheet.getCell(1, 3);
    test:assertEquals(dateCell, "2026-05-28", "Date → ISO string");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCell() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "employees.xlsx");
    Sheet sheet = check wb.getSheet(0);
    // Row 0 is the header row; cell (0,0) = "name"
    CellValue header00 = check sheet.getCell(0, 0);
    test:assertEquals(header00, "name");
    // Row 1, col 0 = first employee's name
    CellValue cell10 = check sheet.getCell(1, 0);
    test:assertEquals(cell10, "John Doe");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellDateTime() returns error? {
    // A datetime cell must bind to its ISO string form with the time component preserved.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    CellValue datetimeCell = check sheet.getCell(1, 4);
    test:assertEquals(datetimeCell, "2026-05-28 14:30:00", "Datetime → ISO string with time");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellBlankIsNil() returns error? {
    // Build a sheet with a known blank cell at (5, 5)
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    CellValue blank = check sheet.getCell(5, 5);
    test:assertEquals(blank, (), "Blank cell should return ()");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellTypedDate() returns error? {
    // Pinning a time:* target binds a date/datetime cell to that record, not the ISO string.
    Workbook wb = check fromFile(TEST_DATA_DIR + "natural_types.xlsx");
    Sheet sheet = check wb.getSheet(0);
    time:Date d = check sheet.getCell(1, 3);
    test:assertEquals(d.year, 2026, "Date cell → time:Date (year)");
    test:assertEquals(d.month, 5, "Date cell → time:Date (month)");
    test:assertEquals(d.day, 28, "Date cell → time:Date (day)");
    time:Civil ts = check sheet.getCell(1, 4);
    test:assertEquals(ts.year, 2026, "Datetime cell → time:Civil (year)");
    test:assertEquals(ts.hour, 14, "Datetime cell → time:Civil (hour)");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetGetCellNonNilablePinOnBlank() returns error? {
    // A non-nilable pinned type over a blank cell must surface a typed error.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    int|Error result = sheet.getCell(5, 5);
    test:assertTrue(result is Error, "Blank cell with a non-nilable target must error");
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
    CellValue v00 = check sheet.getCell(0, 0);
    CellValue v12 = check sheet.getCell(1, 2);
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
    CellValue c00 = check sheet.getCell(0, 0);
    test:assertEquals(c00, 3.14d, "Decimal → decimal");
    CellValue c01 = check sheet.getCell(0, 1);
    test:assertEquals(c01, true, "Boolean → boolean");
    CellValue c02 = check sheet.getCell(0, 2);
    test:assertEquals(c02, "2026-05-28", "Date → ISO string");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetSetCellByAddress() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCellByAddress("A1", "Header");
    check sheet.setCellByAddress("D5", 42.5d);
    CellValue a1 = check sheet.getCell(0, 0);
    CellValue d5 = check sheet.getCell(4, 3);
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
    map<CellValue>[] rows = [
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
    map<CellValue> replacement = {"Name": "Charlie", "Age": 99};
    check sheet.setRow(1, replacement);
    string[][] reread = check sheet.getRows();
    test:assertEquals(reread[0], ["Name", "Age"], "Header row should remain intact after setRow");
    test:assertEquals(reread[1][0], "Charlie", "Row 1 Name should be Charlie");
    test:assertEquals(reread[1][1], "99", "Row 1 Age should be 99");
    check wb.close();
}

// =============================================================================
// putRows / setRow write modes (APPEND / REPLACE / FAIL_IF_EXISTS)
// =============================================================================

@test:Config {groups: ["sheet"]}
function testPutRowsAppendDefault() returns error? {
    // Default mode is APPEND: a second putRows adds below the existing data, not over it.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    check sheet.putRows([["Bob", "25"]]);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 3, "APPEND adds below existing data");
    test:assertEquals(rows[1][0], "Alice", "Existing row preserved");
    test:assertEquals(rows[2][0], "Bob", "New row appended at the bottom");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsReplaceMode() returns error? {
    // REPLACE overwrites from row 0 in place (sheets don't shrink — rows below stay).
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["A", "B"], ["1", "2"], ["3", "4"]]);
    check sheet.putRows([["X", "Y"]], sheetWriteMode = REPLACE);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows[0], ["X", "Y"], "Row 0 overwritten");
    test:assertEquals(rows[1], ["1", "2"], "Rows below are left in place (no shrink)");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsAppendMidInsert() returns error? {
    // APPEND with an explicit startRowIndex inserts there, shifting existing rows down.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["H1", "H2"], ["a", "b"], ["c", "d"]]);
    check sheet.putRows([["X", "Y"]], startRowIndex = 1, sheetWriteMode = APPEND);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 4, "A row was inserted, not overwritten");
    test:assertEquals(rows[1][0], "X", "Inserted at row 1");
    test:assertEquals(rows[2][0], "a", "Existing row shifted down");
    test:assertEquals(rows[3][0], "c", "Trailing row shifted down");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsAppendRecordAlignsToHeader() returns error? {
    // APPEND records below an existing header aligns by column name (field order independent).
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    map<CellValue>[] more = [{"Age": 25, "Name": "Bob"}];
    check sheet.putRows(more);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 3, "Record appended below the existing data");
    test:assertEquals(rows[2][0], "Bob", "Name aligned to column 0 by header");
    test:assertEquals(rows[2][1], "25", "Age aligned to column 1 by header");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsAppendInsertAboveHeaderRefused() returns error? {
    // A record/map APPEND insert at or above the header row corrupts alignment → refused.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    map<CellValue>[] more = [{"Name": "Bob", "Age": 25}];
    Error? result = sheet.putRows(more, startRowIndex = 0, sheetWriteMode = APPEND);
    test:assertTrue(result is Error, "Inserting a record at the header row must be refused");
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 2, "The sheet is left untouched by the refused insert");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsFailIfExists() returns error? {
    // FAIL_IF_EXISTS writes only into empty rows; an occupied target errors and writes nothing.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["A", "B"], ["1", "2"]]);

    Error? hit = sheet.putRows([["X", "Y"]], startRowIndex = 0, sheetWriteMode = FAIL_IF_EXISTS);
    test:assertTrue(hit is Error, "FAIL_IF_EXISTS over an occupied row must error");
    string[][] afterHit = check sheet.getRows();
    test:assertEquals(afterHit[0][0], "A", "Occupied row left untouched after a refused write");

    check sheet.putRows([["X", "Y"]], startRowIndex = 5, sheetWriteMode = FAIL_IF_EXISTS);
    string[][] afterMiss = check sheet.getRows();
    test:assertEquals(afterMiss[5][0], "X", "FAIL_IF_EXISTS writes into an empty row");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSetRowAppendInserts() returns error? {
    // setRow APPEND inserts a row at the index, shifting existing rows down.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["A"], ["B"], ["C"]]);
    check sheet.setRow(1, ["X"], sheetWriteMode = APPEND);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 4, "A row was inserted");
    test:assertEquals(rows[1][0], "X", "Inserted at row 1");
    test:assertEquals(rows[2][0], "B", "Existing row shifted down");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSetRowFailIfExists() returns error? {
    // setRow FAIL_IF_EXISTS errors on an occupied row, writes into an empty one.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["A"], ["B"]]);

    Error? hit = sheet.setRow(0, ["X"], sheetWriteMode = FAIL_IF_EXISTS);
    test:assertTrue(hit is Error, "FAIL_IF_EXISTS over an occupied row must error");

    check sheet.setRow(5, ["Y"], sheetWriteMode = FAIL_IF_EXISTS);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows[0][0], "A", "Occupied row left untouched");
    test:assertEquals(rows[5][0], "Y", "Empty row written");
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

// =============================================================================
// Sheet inserts must not corrupt a table on the same sheet
// =============================================================================

@test:Config {groups: ["sheet"]}
function testPutRowsInsertIntoTableRefused() returns error? {
    // A mid-sheet APPEND whose rows belong to a table would shift the table's cells without its
    // definition following — refuse with a TableOverlapError, leaving the table untouched.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = sheet.putRows([["X", "Y"]], startRowIndex = 1, sheetWriteMode = APPEND);
    test:assertTrue(result is TableOverlapError,
            "Inserting into a table's region via putRows must be refused");

    Table t = check sheet.getTable("T");
    test:assertEquals(check t.getRowCount(), 2, "Table data rows unchanged after a refused insert");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSetRowInsertIntoTableRefused() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = sheet.setRow(1, ["X", "Y"], sheetWriteMode = APPEND);
    test:assertTrue(result is TableOverlapError,
            "Inserting a row into a table's region via setRow must be refused");
    Table t = check sheet.getTable("T");
    test:assertEquals(check t.getRowCount(), 2, "The table is left untouched by the refused insert");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testSheetDeleteRowIntoTableRefused() returns error? {
    // Deleting a sheet row a table sits on would shift its cells without its definition
    // following, so it is refused — Table.deleteRow is the supported path.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = sheet.deleteRow(1);
    test:assertTrue(result is TableOverlapError, "Deleting a row a table sits on must be refused");
    Table t = check sheet.getTable("T");
    test:assertEquals(check t.getRowCount(), 2, "The table is left intact");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsAppendBelowTableSucceeds() returns error? {
    // Appending below a table shifts nothing in the table, so it is allowed.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    check sheet.putRows([["note", "x"]]);
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 3, "Append below the table succeeds");
    test:assertEquals(rows[2][0], "note", "New row landed below the table");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsReplaceIntoTableRegionAllowed() returns error? {
    // REPLACE overwrites in place (no shift), so it leaves the table's ref valid — allowed.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    check sheet.putRows([["Carol", "40"]], startRowIndex = 1, sheetWriteMode = REPLACE);
    Table t = check sheet.getTable("T");
    test:assertEquals(check t.getRowCount(), 2, "Table dimensions unchanged by an in-place REPLACE");
    string[][] rows = check t.getRows();
    test:assertEquals(rows[0][0], "Carol", "Row overwritten in place");
    check wb.close();
}

// =============================================================================
// CODE-REVIEW HARDENING: setRow validate-before-shift + FAIL_IF_EXISTS alignment
// =============================================================================

@test:Config {groups: ["sheet"]}
function testSetRowAppendUnknownFieldLeavesSheetUnchanged() returns error? {
    // An APPEND insert whose record can't be aligned (unknown field) must be rejected BEFORE the
    // row shift, so it leaves no spurious shifted row behind.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["name", "age"], ["Alice", "30"]]);

    record {| string name; int salary; |} bad = {name: "Bob", salary: 5};
    Error? result = sheet.setRow(1, bad, sheetWriteMode = APPEND);
    test:assertTrue(result is Error, "Unknown field must error");
    string[][] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 2, "No spurious row inserted (sheet left unchanged)");
    test:assertEquals(rows[1][0], "Alice", "Existing row not shifted");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsFailIfExistsWritesBelowExistingHeader() returns error? {
    // FAIL_IF_EXISTS with a record aligns to the existing header and writes data below it; an empty
    // data region must be accepted (the header is aligned-to, not an occupied target).
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["name", "age"]]);   // header only, no data rows yet

    map<CellValue>[] rows = [{"name": "Al", "age": 30}];
    check sheet.putRows(rows, sheetWriteMode = FAIL_IF_EXISTS);
    string[][] read = check sheet.getRows();
    test:assertEquals(read.length(), 2, "Header + the appended data row");
    test:assertEquals(read[1][0], "Al", "Aligned to the header and written below it");
    check wb.close();
}

@test:Config {groups: ["sheet"]}
function testPutRowsFailIfExistsRefusesOccupiedDataRow() returns error? {
    // FAIL_IF_EXISTS refuses when a target data cell (resolved by header name) is already occupied.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["name", "age"], ["Existing", "99"]]);

    map<CellValue>[] rows = [{"name": "Al", "age": 30}];
    Error? result = sheet.putRows(rows, sheetWriteMode = FAIL_IF_EXISTS);
    test:assertTrue(result is Error, "An occupied target data row must be refused");
    string[][] read = check sheet.getRows();
    test:assertEquals(read[1][0], "Existing", "Existing data not overwritten");
    check wb.close();
}
