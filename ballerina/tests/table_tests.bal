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
// TABLE TEST DATA SETUP
// =============================================================================

// Record type for table tests
type TableEmployee record {|
    string Name;
    int Age;
    string Department;
|};

// Expected employees from test table
final TableEmployee[] EXPECTED_TABLE_EMPLOYEES = [
    {Name: "Alice", Age: 30, Department: "Engineering"},
    {Name: "Bob", Age: 25, Department: "Marketing"},
    {Name: "Charlie", Age: 35, Department: "Sales"}
];

// Expected string data from test table
final string[][] EXPECTED_TABLE_STRING_DATA = [
    ["Alice", "30", "Engineering"],
    ["Bob", "25", "Marketing"],
    ["Charlie", "35", "Sales"]
];

@test:BeforeSuite
function setupTableTestData() returns error? {
    // Ensure test data directory exists
    if !check file:test(TEST_DATA_DIR, file:EXISTS) {
        check file:createDir(TEST_DATA_DIR, file:RECURSIVE);
    }

    // -------------------------------------------------------------------------
    // tables_test.xlsx - Excel file with tables
    // -------------------------------------------------------------------------
    // Note: We need to create tables programmatically using the Workbook API
    Workbook wb = new;

    // Create a sheet and add some data
    Sheet sheet1 = check wb.createSheet("Employees");
    string[][] employeeData = [
        ["Name", "Age", "Department"],
        ["Alice", "30", "Engineering"],
        ["Bob", "25", "Marketing"],
        ["Charlie", "35", "Sales"]
    ];
    check sheet1.putRows(employeeData);

    // Create a table from the data range
    _ = check sheet1.createTable("EmployeeTable", {
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });

    // Create another sheet with a second table
    Sheet sheet2 = check wb.createSheet("Products");
    string[][] productData = [
        ["Product", "Price", "Stock"],
        ["Widget", "10.99", "100"],
        ["Gadget", "25.50", "50"]
    ];
    check sheet2.putRows(productData);

    _ = check sheet2.createTable("ProductTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });

    check wb.saveAs(TEST_DATA_DIR + "tables_test.xlsx");
    check wb.close();

    // -------------------------------------------------------------------------
    // case_table.xlsx - Table with LOWERCASE headers (name/age/department), read
    // with a record whose fields are capitalised (TableEmployee: Name/Age/Department).
    // Exercises caseInsensitiveHeaders on the table read path.
    // -------------------------------------------------------------------------
    Workbook wbCase = new;
    Sheet caseSheet = check wbCase.createSheet("Data");
    string[][] caseData = [
        ["name", "age", "department"],
        ["Alice", "30", "Engineering"],
        ["Bob", "25", "Marketing"]
    ];
    check caseSheet.putRows(caseData);
    _ = check caseSheet.createTable("CaseTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });
    check wbCase.saveAs(TEST_DATA_DIR + "case_table.xlsx");
    check wbCase.close();
}

@test:AfterSuite
function cleanupTableTestData() returns error? {
    string[] filesToRemove = [
        "tables_test.xlsx",
        "case_table.xlsx",
        "temp_table_write.xlsx",
        "temp_table_create.xlsx",
        "temp_table_write_notfound.xlsx",
        "temp_table_totals.xlsx"
    ];

    foreach string fileName in filesToRemove {
        string filePath = TEST_DATA_DIR + fileName;
        if check file:test(filePath, file:EXISTS) {
            check file:remove(filePath);
        }
    }
}

// =============================================================================
// SIMPLE TABLE API TESTS (parseTable/writeTable)
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testParseTableToRecords() returns error? {
    TableEmployee[] employees = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable");

    test:assertEquals(employees.length(), 3, "Should have 3 employees");
    test:assertEquals(employees[0].Name, "Alice", "First employee name");
    test:assertEquals(employees[0].Age, 30, "First employee age");
    test:assertEquals(employees[0].Department, "Engineering", "First employee department");
    test:assertEquals(employees[2].Name, "Charlie", "Third employee name");
}

@test:Config {
    groups: ["table"]
}
function testParseTableToStringArray() returns error? {
    string[][] data = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable");

    assertStringArrayEquals(data, EXPECTED_TABLE_STRING_DATA, "Table string rows");
}

@test:Config {
    groups: ["table"]
}
function testParseTableToMaps() returns error? {
    map<CellValue>[] data = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable");

    test:assertEquals(data.length(), 3, "Should have 3 maps");
    test:assertEquals(data[0]["Name"], "Alice", "First employee name");
    test:assertEquals(data[0]["Department"], "Engineering", "First employee department");
}

@test:Config {
    groups: ["table"]
}
function testParseTableNotFound() returns error? {
    TableEmployee[]|Error result = parseTable(TEST_DATA_DIR + "tables_test.xlsx", "NonExistentTable");

    test:assertTrue(result is TableNotFoundError, "Should return TableNotFoundError");
    if result is TableNotFoundError {
        test:assertTrue(result.message().includes("NonExistentTable"), "Error should mention table name");
    }
}

@test:Config {
    groups: ["table"]
}
function testParseTableCaseInsensitiveHeaders() returns error? {
    // CaseTable has lowercase headers (name/age/department); TableEmployee fields are
    // capitalised (Name/Age/Department). With caseInsensitiveHeaders they must match.
    TableParseOptions opts = {
        caseInsensitiveHeaders: true
    };

    TableEmployee[] employees = check parseTable(TEST_DATA_DIR + "case_table.xlsx", "CaseTable", opts);

    test:assertEquals(employees.length(), 2, "Should have 2 employees");
    test:assertEquals(employees[0].Name, "Alice", "First employee name");
    test:assertEquals(employees[0].Age, 30, "First employee age");
    test:assertEquals(employees[1].Name, "Bob", "Second employee name");
}

@test:Config {
    groups: ["table"]
}
function testTableGetRowWithOptions() returns error? {
    // Single-row table read threads its binding fields through TableRowParseOptions.
    // CaseTable has lowercase headers (name/age) vs the capitalised TableEmployee fields,
    // so the row only binds when caseInsensitiveHeaders is honoured by getRow.
    Workbook wb = check fromFile(TEST_DATA_DIR + "case_table.xlsx");
    Table tbl = check wb.getTable("CaseTable");

    TableRowParseOptions opts = {
        caseInsensitiveHeaders: true
    };
    TableEmployee emp = check tbl.getRow(0, opts);

    test:assertEquals(emp.Name, "Alice", "getRow(0) name via caseInsensitive match");
    test:assertEquals(emp.Age, 30, "getRow(0) age via caseInsensitive match");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWriteTableExpands() returns error? {
    // First, copy the test file
    string tempFile = TEST_DATA_DIR + "temp_table_write.xlsx";

    // Open original and save as temp
    Workbook wbSrc = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    check wbSrc.saveAs(tempFile);
    check wbSrc.close();

    // Write more data to the table (should auto-expand)
    TableEmployee[] newEmployees = [
        {Name: "David", Age: 28, Department: "HR"},
        {Name: "Eve", Age: 32, Department: "Finance"},
        {Name: "Frank", Age: 40, Department: "Legal"},
        {Name: "Grace", Age: 29, Department: "IT"}
    ];
    check writeTable(newEmployees, tempFile, "EmployeeTable");

    // Read back and verify
    TableEmployee[] result = check parseTable(tempFile, "EmployeeTable");
    test:assertEquals(result.length(), 4, "Should have 4 employees after write");
    test:assertEquals(result[0].Name, "David", "First employee should be David");
    test:assertEquals(result[3].Name, "Grace", "Last employee should be Grace");

    // Verify the table RANGE metadata actually expanded (original had 3 data rows),
    // not just that the data was written.
    Workbook wbCheck = check fromFile(tempFile);
    Table expanded = check wbCheck.getTable("EmployeeTable");
    test:assertEquals(check expanded.getRowCount(), 4, "Table data-row count should be 4 after auto-expand");
    CellRange dataRange = check expanded.getDataCellRange();
    test:assertEquals(dataRange.lastRowIndex, 4, "Data range should extend to row 4 after auto-expand");
    check wbCheck.close();
}

// =============================================================================
// WORKBOOK TABLE API TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testWorkbookGetTable() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");

    Table empTable = check wb.getTable("EmployeeTable");
    test:assertEquals(check empTable.getName(), "EmployeeTable", "Table name should match");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWorkbookGetTableNotFound() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");

    Table|Error result = wb.getTable("NonExistentTable");
    test:assertTrue(result is TableNotFoundError, "Should return TableNotFoundError");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWorkbookGetAllTables() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");

    Table[] tables = check wb.getAllTables();
    test:assertEquals(tables.length(), 2, "Should have 2 tables");

    // Tables could be in any order
    string[] tableNames = [];
    foreach Table t in tables {
        tableNames.push(check t.getName());
    }
    test:assertTrue(tableNames.indexOf("EmployeeTable") != (), "Should contain EmployeeTable");
    test:assertTrue(tableNames.indexOf("ProductTable") != (), "Should contain ProductTable");

    check wb.close();
}

// =============================================================================
// SHEET TABLE API TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testSheetGetTable() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table empTable = check sheet.getTable("EmployeeTable");
    test:assertEquals(check empTable.getName(), "EmployeeTable", "Table name should match");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetGetTables() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table[] tables = check sheet.getTables();
    test:assertEquals(tables.length(), 1, "Employees sheet should have 1 table");
    test:assertEquals(check tables[0].getName(), "EmployeeTable", "Table name should be EmployeeTable");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetCreateTable() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    // First write some data
    string[][] data = [
        ["Col1", "Col2"],
        ["A", "B"],
        ["C", "D"]
    ];
    check sheet.putRows(data);

    // Create a table from the data range
    Table newTable = check sheet.createTable("NewTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    test:assertEquals(check newTable.getName(), "NewTable", "Table name should match");
    test:assertEquals(check newTable.getRowCount(), 2, "Should have 2 data rows");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetCreateTableFromData() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    TableEmployee[] employees = [
        {Name: "Test1", Age: 20, Department: "Dept1"},
        {Name: "Test2", Age: 30, Department: "Dept2"}
    ];

    Table newTable = check sheet.createTableFromData("EmpTable", employees);

    test:assertEquals(check newTable.getName(), "EmpTable", "Table name should match");
    test:assertEquals(check newTable.getColumnCount(), 3, "Should have 3 columns");
    test:assertEquals(check newTable.getRowCount(), 2, "Should have 2 data rows");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testCreateTableFromDataWithStringArray() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    // For string[][] the first row is the header. The table height must equal the
    // supplied row count (3) — header at row 0, data at rows 1-2 — with no phantom
    // trailing row (regression test for the createTableFromData off-by-one).
    string[][] data = [
        ["Col1", "Col2"],
        ["A", "B"],
        ["C", "D"]
    ];

    Table t = check sheet.createTableFromData("StringTable", data);
    test:assertEquals(check t.getRowCount(), 2, "Should have exactly 2 data rows, no phantom row");
    CellRange range = check t.getCellRange();
    test:assertEquals(range.lastRowIndex, 2, "Last row index should equal data height (header 0, data 1-2)");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetDeleteTable() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    // Create data and table
    string[][] data = [["A", "B"], ["1", "2"]];
    check sheet.putRows(data);
    _ = check sheet.createTable("TempTable", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // Verify table exists
    Table[] tablesBefore = check sheet.getTables();
    test:assertEquals(tablesBefore.length(), 1, "Should have 1 table before delete");

    // Delete the table
    check sheet.deleteTable("TempTable");

    // Verify table is deleted
    Table[] tablesAfter = check sheet.getTables();
    test:assertEquals(tablesAfter.length(), 0, "Should have 0 tables after delete");

    check wb.close();
}

// =============================================================================
// TABLE CLASS TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableIdentityMethods() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    test:assertEquals(check empTable.getName(), "EmployeeTable", "getName() should return table name");
    test:assertEquals(check empTable.getDisplayName(), "EmployeeTable", "getDisplayName() should return display name");
    test:assertEquals(check empTable.getSheetName(), "Employees", "getSheetName() should return sheet name");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableRangeMethods() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    // Full range (including header)
    CellRange fullRange = check empTable.getCellRange();
    test:assertEquals(fullRange.firstRowIndex, 0, "Full range starts at row 0");
    test:assertEquals(fullRange.lastRowIndex, 3, "Full range ends at row 3");

    // Data range (excluding header)
    CellRange dataRange = check empTable.getDataCellRange();
    test:assertEquals(dataRange.firstRowIndex, 1, "Data range starts at row 1");
    test:assertEquals(dataRange.lastRowIndex, 3, "Data range ends at row 3");

    test:assertEquals(check empTable.getRowCount(), 3, "Should have 3 data rows");
    test:assertEquals(check empTable.getColumnCount(), 3, "Should have 3 columns");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetHeaders() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    string[] headers = check empTable.getHeaders();
    test:assertEquals(headers.length(), 3, "Should have 3 headers");
    test:assertEquals(headers[0], "Name", "First header should be Name");
    test:assertEquals(headers[1], "Age", "Second header should be Age");
    test:assertEquals(headers[2], "Department", "Third header should be Department");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetRows() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    TableEmployee[] employees = check empTable.getRows();

    test:assertEquals(employees, EXPECTED_TABLE_EMPLOYEES, "Table rows should match expected employees");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetRow() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    TableEmployee emp = check empTable.getRow(1);
    test:assertEquals(emp.Name, "Bob", "Second employee (index 1) should be Bob");
    test:assertEquals(emp.Age, 25, "Bob's age should be 25");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTablePutRows() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Test");

    // Create initial data and table
    string[][] initialData = [
        ["Name", "Value"],
        ["A", "1"],
        ["B", "2"]
    ];
    check sheet.putRows(initialData);

    Table t = check sheet.createTable("TestTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // Put new rows
    string[][] newData = [
        ["X", "10"],
        ["Y", "20"],
        ["Z", "30"]
    ];
    check t.putRows(newData);

    // Verify
    string[][] result = check t.getRows();
    test:assertEquals(result.length(), 3, "Should have 3 rows");
    test:assertEquals(result[0][0], "X", "First value should be X");
    test:assertEquals(result[2][1], "30", "Last value should be 30");

    check wb.close();
}

// =============================================================================
// TABLE WRITE-MODE TESTS — REPLACE resize (grow/shrink/clear), APPEND, overlap
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testWriteTableReplaceShrinks() returns error? {
    // REPLACE (the default) with fewer rows must shrink the data range — no stale rows survive
    // inside the table. EmployeeTable starts with 3 data rows.
    string tempFile = getTempFilePath("table_shrink");
    Workbook wbSrc = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    check wbSrc.saveAs(tempFile);
    check wbSrc.close();

    TableEmployee[] fewer = [{Name: "Solo", Age: 41, Department: "Ops"}];
    check writeTable(fewer, tempFile, "EmployeeTable");

    TableEmployee[] result = check parseTable(tempFile, "EmployeeTable");
    test:assertEquals(result.length(), 1, "REPLACE with 1 row must leave exactly 1 data row");
    test:assertEquals(result[0].Name, "Solo", "The single row must be the written row");

    Workbook wbCheck = check fromFile(tempFile);
    Table shrunk = check wbCheck.getTable("EmployeeTable");
    test:assertEquals(check shrunk.getRowCount(), 1, "Table data-row count must shrink to 1");
    CellRange dataRange = check shrunk.getDataCellRange();
    test:assertEquals(dataRange.lastRowIndex, 1, "Data range must shrink (no stale rows below)");
    check wbCheck.close();
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["table"]
}
function testWriteTableReplaceClears() returns error? {
    // REPLACE with an empty array clears the table to a single blank data row (Excel needs >= 1).
    string tempFile = getTempFilePath("table_clear");
    Workbook wbSrc = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    check wbSrc.saveAs(tempFile);
    check wbSrc.close();

    TableEmployee[] none = [];
    check writeTable(none, tempFile, "EmployeeTable");

    Workbook wbCheck = check fromFile(tempFile);
    Table cleared = check wbCheck.getTable("EmployeeTable");
    test:assertEquals(check cleared.getRowCount(), 1, "Cleared table keeps a single blank data row");
    string[][] rows = check cleared.getRows();
    test:assertEquals(rows.length(), 1, "Cleared table reads a single data row");
    test:assertEquals(rows[0][0], "", "The kept data row must be blank");
    check wbCheck.close();
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["table"]
}
function testWriteTableAppendMode() returns error? {
    // APPEND adds rows below the existing data; existing rows are preserved.
    string tempFile = getTempFilePath("table_append");
    Workbook wbSrc = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    check wbSrc.saveAs(tempFile);
    check wbSrc.close();

    TableEmployee[] more = [
        {Name: "Dan", Age: 28, Department: "HR"},
        {Name: "Eve", Age: 33, Department: "Legal"}
    ];
    check writeTable(more, tempFile, "EmployeeTable", tableWriteMode = APPEND);

    TableEmployee[] result = check parseTable(tempFile, "EmployeeTable");
    test:assertEquals(result.length(), 5, "APPEND must keep 3 original + 2 new rows");
    test:assertEquals(result[0].Name, "Alice", "Original first row preserved");
    test:assertEquals(result[2].Name, "Charlie", "Original last row preserved");
    test:assertEquals(result[3].Name, "Dan", "First appended row");
    test:assertEquals(result[4].Name, "Eve", "Second appended row");
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["table"]
}
function testTablePutRowsAppendMode() returns error? {
    // The object API honours the mode too.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("T");
    check sheet.putRows([["Name", "Value"], ["A", "1"], ["B", "2"]]);
    Table t = check sheet.createTable("PutAppend", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    check t.putRows([["C", "3"]], tableWriteMode = APPEND);

    string[][] rows = check t.getRows();
    test:assertEquals(rows.length(), 3, "APPEND adds below existing data");
    test:assertEquals(rows[0][0], "A", "Original row preserved");
    test:assertEquals(rows[2][0], "C", "Appended row at the bottom");
    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWriteTableTotalsRowResize() returns error? {
    // A totals row must survive a resize (grow then shrink) and stay directly below the data.
    string totalsFile = getTempFilePath("table_totals_resize");
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Sales");
    check sheet.putRows([["Region", "Amount"], ["North", "100"], ["South", "250"]]);
    Table salesTable = check sheet.createTable("SalesTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    check setTotalRowNative(salesTable, 1, 350.0);
    check wb.saveAs(totalsFile);
    check wb.close();

    // Grow: replace 2 data rows with 4.
    string[][] grown = [["North", "100"], ["South", "250"], ["East", "75"], ["West", "125"]];
    check writeTable(grown, totalsFile, "SalesTable");

    Workbook wbGrown = check fromFile(totalsFile);
    Table grownTable = check wbGrown.getTable("SalesTable");
    test:assertEquals(check grownTable.getRowCount(), 4, "Data rows grow to 4");
    test:assertTrue(check grownTable.hasTotalRow(), "Totals row survives the grow");
    map<CellValue> grownTotals = check grownTable.getTotalRow();
    test:assertEquals(grownTotals["Amount"], 350, "Totals value carried by the shift");
    check wbGrown.close();

    // Shrink: replace 4 data rows with 1.
    string[][] shrunk = [["Only", "999"]];
    check writeTable(shrunk, totalsFile, "SalesTable");

    Workbook wbShrunk = check fromFile(totalsFile);
    Table shrunkTable = check wbShrunk.getTable("SalesTable");
    test:assertEquals(check shrunkTable.getRowCount(), 1, "Data rows shrink to 1 (no stale rows)");
    test:assertTrue(check shrunkTable.hasTotalRow(), "Totals row survives the shrink");
    map<CellValue> shrunkTotals = check shrunkTable.getTotalRow();
    test:assertEquals(shrunkTotals["Amount"], 350, "Totals value still intact after shrink");
    check wbShrunk.close();
    check removeTempFile(totalsFile);
}

@test:Config {
    groups: ["table"]
}
function testWriteTableResizeOverlapError() returns error? {
    // Two tables stacked on one sheet. Growing the upper one over the lower must fail loud with a
    // TableOverlapError and leave both tables untouched.
    string tempFile = getTempFilePath("table_overlap");
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Stacked");
    check sheet.putRows([
        ["Name", "Value"],
        ["A", "1"],
        ["B", "2"],
        ["", ""],
        ["P", "Q"],
        ["x", "y"],
        ["z", "w"]
    ]);
    _ = check sheet.createTable("Upper", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    _ = check sheet.createTable("Lower", {
        firstRowIndex: 4,
        lastRowIndex: 6,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    check wb.saveAs(tempFile);
    check wb.close();

    // Grow Upper from 2 to 5 data rows — would shift Lower down → overlap.
    string[][] big = [["A", "1"], ["B", "2"], ["C", "3"], ["D", "4"], ["E", "5"]];
    Error? result = writeTable(big, tempFile, "Upper");
    test:assertTrue(result is TableOverlapError,
            "Growing a table into another must return TableOverlapError");

    // Both tables untouched — nothing was written.
    Workbook wbCheck = check fromFile(tempFile);
    Table upperCheck = check wbCheck.getTable("Upper");
    Table lowerCheck = check wbCheck.getTable("Lower");
    test:assertEquals(check upperCheck.getRowCount(), 2, "Upper table unchanged after a refused write");
    test:assertEquals(check lowerCheck.getRowCount(), 2, "Lower table unchanged after a refused write");
    check wbCheck.close();
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["table"]
}
function testWriteTableShiftsPlainContentDown() returns error? {
    // Plain (non-table) content below a table is carried down by a grow, not overwritten.
    string tempFile = getTempFilePath("table_plain_below");
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([
        ["Name", "Value"],
        ["A", "1"],
        ["B", "2"],
        ["", ""],
        ["NOTE", "keep-me"]
    ]);
    _ = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    check wb.saveAs(tempFile);
    check wb.close();

    // Grow the table by 2 rows; the note at row 4 must shift down to row 6, preserved.
    string[][] more = [["A", "1"], ["B", "2"], ["C", "3"], ["D", "4"]];
    check writeTable(more, tempFile, "T");

    Workbook wbCheck = check fromFile(tempFile);
    Sheet s = check wbCheck.getSheet("S");
    string moved = check s.getCell(6, 1);
    test:assertEquals(moved, "keep-me", "Plain content below the table is carried down, not clobbered");
    check wbCheck.close();
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["table"]
}
function testTableRename() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Test");

    // Create a table
    string[][] data = [["A", "B"], ["1", "2"]];
    check sheet.putRows(data);
    Table t = check sheet.createTable("OldName", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    test:assertEquals(check t.getName(), "OldName", "Initial name should be OldName");

    // Rename
    check t.rename("NewName");
    test:assertEquals(check t.getName(), "NewName", "Name should be NewName after rename");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableResize() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Test");

    // Create initial data (more than the table will initially cover)
    string[][] data = [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["4", "5", "6"],
        ["7", "8", "9"]
    ];
    check sheet.putRows(data);

    // Create table covering only first 2 rows
    Table t = check sheet.createTable("ResizeTable", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    test:assertEquals(check t.getRowCount(), 1, "Initial row count should be 1");

    // Resize to include more data
    check t.resize({
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });

    test:assertEquals(check t.getRowCount(), 3, "Row count should be 3 after resize");
    test:assertEquals(check t.getColumnCount(), 3, "Column count should be 3 after resize");

    check wb.close();
}

// =============================================================================
// ERROR HANDLING TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableNotFoundErrorType() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");

    Table|Error result = wb.getTable("NonExistent");
    test:assertTrue(result is TableNotFoundError, "Should be TableNotFoundError");

    if result is TableNotFoundError {
        ErrorDetails? details = result.detail();
        test:assertTrue(details?.tableName == "NonExistent", "Error details should contain table name");
    }

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetTableNotFoundError() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table|Error result = sheet.getTable("NonExistent");
    test:assertTrue(result is TableNotFoundError, "Should be TableNotFoundError");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testDeleteNonExistentTable() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Test");

    Error? result = sheet.deleteTable("NonExistent");
    test:assertTrue(result is TableNotFoundError, "Should return TableNotFoundError");

    check wb.close();
}

// =============================================================================
// TOTALS ROW TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableHasTotalsRowFalse() returns error? {
    // Our programmatically created tables don't have totals rows
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    test:assertFalse(check empTable.hasTotalRow(), "Table should not have totals row");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetTotalsRowError() returns error? {
    // Getting totals row from table without totals should return error
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    map<CellValue>|Error result = empTable.getTotalRow();
    test:assertTrue(result is Error, "Should return error when table has no totals row");
    if result is Error {
        test:assertTrue(result.message().includes("does not have a total row"),
                "Error must identify the missing-total-row cause");
    }

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetTotalRow() returns error? {
    // Build a table, author a total row via the test-only helper, then read it back.
    // The total-row map must come back typed as map<CellValue> with the total bound
    // to its natural Ballerina type (a whole number → int).
    string totalsFile = TEST_DATA_DIR + "temp_table_totals.xlsx";
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Sales");
    string[][] salesData = [
        ["Region", "Amount"],
        ["North", "100"],
        ["South", "250"]
    ];
    check sheet.putRows(salesData);
    Table salesTable = check sheet.createTable("SalesTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    // Write 350 as the total into the "Amount" column (table column index 1).
    check setTotalRowNative(salesTable, 1, 350.0);
    check wb.saveAs(totalsFile);
    check wb.close();

    Workbook reopened = check fromFile(totalsFile);
    Table reopenedTable = check reopened.getTable("SalesTable");

    test:assertTrue(check reopenedTable.hasTotalRow(), "Table should have a total row");

    map<CellValue> totals = check reopenedTable.getTotalRow();
    // The returned map must be a genuine map<CellValue>, not the wider map<anydata> the
    // native builds internally — this guards the typedesc-based inherent-type fix.
    test:assertTrue(totals is map<CellValue>,
            "Total row must be a genuine map<CellValue>, not map<anydata>");
    CellValue amountTotal = totals["Amount"];
    test:assertEquals(amountTotal, 350, "Total should bind to its natural type (int 350)");
    test:assertTrue(amountTotal is int, "Whole-number total must bind to int, not decimal/string");

    check reopened.close();
}

// =============================================================================
// ADDITIONAL ERROR SCENARIO TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableOverlapError() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    // Create initial data
    string[][] data = [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["4", "5", "6"]
    ];
    check sheet.putRows(data);

    // Create first table
    _ = check sheet.createTable("FirstTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });

    // Try to create overlapping table - should fail
    Table|Error result = sheet.createTable("OverlappingTable", {
        firstRowIndex: 1,
        lastRowIndex: 3,
        firstColumnIndex: 1,
        lastColumnIndex: 3
    });

    test:assertTrue(result is TableOverlapError, "Should return TableOverlapError for overlapping tables");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testInvalidTableRangeError() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("TestSheet");

    string[][] data = [["A", "B"], ["1", "2"], ["3", "4"]];
    check sheet.putRows(data);

    Table t = check sheet.createTable("TestTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // Resizing to a header-only range (firstRow == lastRow, no data row) is rejected:
    // an Excel table must keep at least one header row and one data row.
    Error? result = t.resize({
        firstRowIndex: 0,
        lastRowIndex: 0,  // Only header row, no data row
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    test:assertTrue(result is InvalidTableRangeError,
            "Resizing to a header-only range must return InvalidTableRangeError");
    if result is InvalidTableRangeError {
        test:assertTrue(result.message().includes("at least one header row and one data row"),
                "Error must identify the missing data row");
    }

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testGetRowOutOfRange() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    // Table has 3 data rows (indices 0, 1, 2)
    // Try to get row at index 10 (out of range)
    TableEmployee|Error result = empTable.getRow(10);
    test:assertTrue(result is Error, "Should return error for out of range index");

    // Also test negative index
    TableEmployee|Error negResult = empTable.getRow(-1);
    test:assertTrue(negResult is Error, "Should return error for negative index");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWriteTableNotFound() returns error? {
    string tempFile = TEST_DATA_DIR + "temp_table_write_notfound.xlsx";

    // Create a file without the target table
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Sheet1");
    string[][] data = [["A", "B"], ["1", "2"]];
    check sheet.putRows(data);
    check wb.saveAs(tempFile);
    check wb.close();

    // Try to write to non-existent table
    TableEmployee[] employees = [{Name: "Test", Age: 20, Department: "Dept"}];
    Error? result = writeTable(employees, tempFile, "NonExistentTable");

    test:assertTrue(result is TableNotFoundError, "Should return TableNotFoundError");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// OPTIONS TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testParseTableWithRowCount() returns error? {
    // Test that rowCount option limits the number of rows returned
    TableParseOptions opts = {
        rowCount: 2
    };
    TableEmployee[] employees = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable", opts);

    // EmployeeTable has 3 data rows, but we limit to 2
    test:assertEquals(employees.length(), 2, "Should return only 2 rows due to rowCount limit");
    test:assertEquals(employees[0].Name, "Alice", "First employee should be Alice");
    test:assertEquals(employees[1].Name, "Bob", "Second employee should be Bob");
}

@test:Config {
    groups: ["table"]
}
function testTableGetRowsWithRowCount() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    // Get rows with rowCount limit
    TableParseOptions opts = {
        rowCount: 1
    };
    TableEmployee[] employees = check empTable.getRows(opts);

    test:assertEquals(employees.length(), 1, "Should return only 1 row due to rowCount limit");
    test:assertEquals(employees[0].Name, "Alice", "First employee should be Alice");

    check wb.close();
}

// =============================================================================
// Public Data/Row union dispatch — Table API surface
// =============================================================================

@test:Config {groups: ["table"]}
function testWriteTableWithInlineLiteral() returns error? {
    // Set up: copy tables_test.xlsx to a temp location so we don't mutate the shared fixture
    string sourceFile = TEST_DATA_DIR + "tables_test.xlsx";
    string tempFile = getTempFilePath("inline_writetable");
    Workbook srcWb = check fromFile(sourceFile);
    byte[] bytes = check srcWb.toBytes();
    check srcWb.close();
    Workbook destWb = check fromBytes(bytes);
    check destWb.saveAs(tempFile);
    check destWb.close();

    // Inline literal — contextually typed against `Data = Row[]`. Exercises the
    // table writer's row-by-row dispatch (writeRowData inspects runtime types).
    check writeTable([["Eve", "40", "Finance"], ["Frank", "33", "HR"]], tempFile,
            "EmployeeTable");

    Workbook check_wb = check fromFile(tempFile);
    Table empTable = check check_wb.getTable("EmployeeTable");
    string[][] rows = check empTable.getRows();
    // REPLACE (the default) resizes the table to fit the data, so writing two rows over the
    // original three shrinks it to two — no stale trailing row survives.
    test:assertEquals(rows.length(), 2, "REPLACE resizes the table to fit (3 -> 2 rows, no stale row)");
    test:assertEquals(rows[0][0], "Eve");
    test:assertEquals(rows[1][0], "Frank");
    check check_wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["table"]}
function testTableGetRowsWithDataTarget() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");
    // Explicit `Row[]` target — the inferred element is the `Row` union itself;
    // dispatch must handle the UNION element tag and fall back to string[][].
    Row[] rows = check empTable.getRows();
    test:assertTrue(rows is string[][], "Row[] target on Table.getRows should fall back to string[][]");
    if rows is string[][] {
        test:assertEquals(rows.length(), 3, "EmployeeTable has 3 data rows");
    }
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableGetRowWithRowTarget() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");
    Row r = check empTable.getRow(0);
    test:assertTrue(r is string[], "Row target on Table.getRow should fall back to string[]");
    if r is string[] {
        test:assertEquals(r[0], "Alice");
    }
    check wb.close();
}

@test:Config {groups: ["table"]}
function testUseTableAfterCloseReturnsError() returns error? {
    Workbook wb = check fromFile(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");
    check wb.close();
    string[][]|Error result = empTable.getRows();
    test:assertTrue(result is Error, "Table.getRows after Workbook.close should return Error");
    if result is Error {
        test:assertTrue(result.message().includes("no longer valid"),
                "Error must indicate the handle was invalidated");
    }
}

// =============================================================================
// Table name validation
// =============================================================================

@test:Config {groups: ["table"]}
function testCreateTableInvalidNameWithSpace() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    Table|Error result = sheet.createTable("Bad Table", "A1:B2");
    test:assertTrue(result is Error, "Table name with space must be rejected");
    if result is Error {
        test:assertTrue(result.message().includes("cannot contain spaces"),
                "Error must identify the space as the cause");
    }
    check wb.close();
}

@test:Config {groups: ["table"]}
function testCreateTableNameStartsWithDigit() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    Table|Error result = sheet.createTable("1Sales", "A1:B2");
    test:assertTrue(result is Error,
            "Table name starting with a digit must be rejected");
    if result is Error {
        test:assertTrue(result.message().includes("must start with a letter or underscore"),
                "Error must identify the leading-digit as the cause");
    }
    check wb.close();
}

@test:Config {groups: ["table"]}
function testWriteTableFailureDoesNotLeak() returns error? {
    string tempFile = TEST_DATA_DIR + "table_leak_regression.xlsx";

    Workbook wb = new;
    Sheet sh = check wb.createSheet("Data");
    check sh.putRows([["msg"], ["init"]]);
    _ = check sh.createTable("Logs", "A1:A2");
    check wb.saveAs(tempFile);
    check wb.close();

    // Excel's per-cell limit is 32_767 chars. Build a 65_536-char string by
    // doubling "x" 16 times — POI rejects this during write.
    string oversized = "x";
    foreach int _ in 0 ..< 16 {
        oversized = oversized + oversized;
    }
    record {|string msg;|}[] badRow = [{msg: oversized}];

    foreach int _ in 0 ..< 50 {
        Error? r = writeTable(badRow, tempFile, "Logs");
        test:assertTrue(r is Error,
                "Oversized writeTable must return an Error (not panic, not silent success)");
    }

    // After 50 failed writes, a clean write must still succeed AND persist.
    record {|string msg;|}[] goodRow = [{msg: "after-failures"}];
    check writeTable(goodRow, tempFile, "Logs");

    // parseTable returns data rows only (no header). Robust check: look for
    // the marker value rather than asserting an exact row count, so the test
    // doesn't depend on writeTable's replace-vs-append semantics.
    string[][] parsed = check parseTable(tempFile, "Logs");
    boolean foundMarker = false;
    foreach string[] row in parsed {
        if row.length() > 0 && row[0] == "after-failures" {
            foundMarker = true;
            break;
        }
    }
    test:assertTrue(foundMarker,
            "Clean writeTable after 50 cleanup cycles must persist its data");

    check file:remove(tempFile);
}

// =============================================================================
// createTableFromData honours startColumnIndex
// =============================================================================
// The startColumnIndex parameter on Sheet.createTableFromData used to be
// silently ignored on write — the table area was created at the requested
// column but data landed at column 0. The fix threads startColumnIndex
// through XlsxWriter so both the table metadata and the cells agree.

type OffsetTableRow record {|
    string name;
    int age;
|};

@test:Config {groups: ["table"]}
function testCreateTableFromDataAtNonZeroStartColumn() returns error? {
    string tempFile = getTempFilePath("table_offset_col");

    OffsetTableRow[] data = [
        {name: "Alice", age: 30},
        {name: "Bob", age: 25}
    ];

    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    _ = check sheet.createTableFromData("OffsetTable", data, 2, 3);
    check wb.saveAs(tempFile);
    check wb.close();

    // Re-open and verify the table is at the requested offset (row 2, col 3),
    // and that the data cells actually landed there — not at column 0.
    Workbook wb2 = check fromFile(tempFile);
    Table t2 = check wb2.getTable("OffsetTable");
    CellRange range = check t2.getCellRange();
    test:assertEquals(range.firstRowIndex, 2);
    test:assertEquals(range.firstColumnIndex, 3);
    test:assertEquals(range.lastColumnIndex, 4);

    Sheet s2 = check wb2.getSheet(0);
    CellValue headerNameCell = check s2.getCell(2, 3);
    test:assertEquals(headerNameCell, "name",
            "Header 'name' must be at row 2 col 3, matching the requested table offset");
    CellValue firstNameCell = check s2.getCell(3, 3);
    test:assertEquals(firstNameCell, "Alice",
            "Data value 'Alice' must be at row 3 col 3, not column 0");
    check wb2.close();
    check file:remove(tempFile);
}

// =============================================================================
// Table writes resolve columns by header, not key order
// =============================================================================
// TableHandle.writeRowData used to place values positionally by key/field
// iteration order, which silently misaligned data when the record/map shape
// didn't match the table's column declaration order. The fix looks each
// header up by name (with @xlsx:Name support for records).

type AgeFirstRow record {|
    int age;
    string name;
|};

type NameFirstRow record {|
    string name;
    int age;
|};

@test:Config {groups: ["table"]}
function testTableWriteResolvesColumnsByHeader() returns error? {
    string tempFile = getTempFilePath("table_header_lookup");

    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    // Headers declared in order: name, age.
    check sheet.putRows([["name", "age"], ["seed", "0"]]);
    _ = check sheet.createTable("Employees", "A1:B2");
    check wb.saveAs(tempFile);
    check wb.close();

    // Write a record whose FIELD declaration order is age, name — reversed
    // from the table's header order. Positional writing (the old behaviour)
    // would put age in column 0 and name in column 1; header-resolved writing
    // (the fix) places them correctly regardless of declaration order.
    AgeFirstRow[] reverseOrderRow = [{age: 42, name: "Alice"}];
    check writeTable(reverseOrderRow, tempFile, "Employees");

    // Round-trip the table back through parseTable with a header-aligned
    // record. parseTable maps headers → fields by name, so if writeTable had
    // placed the values in the wrong columns, name would receive an integer
    // (TypeConversionError) or age would receive a string ("Alice", also a
    // TypeConversionError). A clean round-trip is the proof.
    NameFirstRow[] parsed = check parseTable(tempFile, "Employees");
    test:assertEquals(parsed.length(), 1,
            "Table should contain exactly one data row after the overwrite");
    test:assertEquals(parsed[0].name, "Alice",
            "name field must land in the 'name' column, regardless of write key order");
    test:assertEquals(parsed[0].age, 42,
            "age field must land in the 'age' column, regardless of write key order");
    check file:remove(tempFile);
}

// =============================================================================
// TABLE MID-POSITION INSERT (APPEND insertAt)
// =============================================================================

@test:Config {groups: ["table"]}
function testTablePutRowsInsertAt() returns error? {
    // APPEND with insertAt inserts inside the table, shifting existing rows down.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"], ["Cara", "40"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // Insert at data-row 1 (between Alice and Bob).
    check t.putRows([["Zoe", "99"]], tableWriteMode = APPEND, insertAt = 1);

    string[][] rows = check t.getRows();
    test:assertEquals(rows.length(), 4, "Table grew by one row");
    test:assertEquals(rows[0][0], "Alice", "Row before the insert preserved");
    test:assertEquals(rows[1][0], "Zoe", "Inserted at data-row 1");
    test:assertEquals(rows[2][0], "Bob", "Existing row shifted down");
    test:assertEquals(rows[3][0], "Cara", "Trailing row shifted down");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableInsertAtOutOfRange() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // currentDataRows = 1; insertAt = 5 is out of range.
    Error? result = t.putRows([["X", "9"]], tableWriteMode = APPEND, insertAt = 5);
    test:assertTrue(result is InvalidTableRangeError, "insertAt out of range must error");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableInsertAtIgnoredForReplace() returns error? {
    // insertAt is APPEND-only; with REPLACE it is ignored (the whole data region is replaced).
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    check t.putRows([["Solo", "1"]], tableWriteMode = REPLACE, insertAt = 1);
    string[][] rows = check t.getRows();
    test:assertEquals(rows.length(), 1, "REPLACE replaces the whole data region, ignoring insertAt");
    test:assertEquals(rows[0][0], "Solo", "Data replaced");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableInsertAtOverlapError() returns error? {
    // A mid-insert that would shift a second table below must error.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([
        ["Name", "Age"],
        ["Alice", "30"],
        ["Bob", "25"],
        ["", ""],
        ["P", "Q"],
        ["x", "y"]
    ]);
    Table t = check sheet.createTable("Upper", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    _ = check sheet.createTable("Lower", {
        firstRowIndex: 4,
        lastRowIndex: 5,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = t.putRows([["Z", "9"]], tableWriteMode = APPEND, insertAt = 0);
    test:assertTrue(result is TableOverlapError, "A mid-insert that shifts a second table must error");
    check wb.close();
}

// =============================================================================
// TABLE ROW DELETE (Table.deleteRow)
// =============================================================================

// Record whose optional field may be absent in a given row.
type OptionalAgeRow record {|
    string name;
    int age?;
|};

@test:Config {groups: ["table"]}
function testTableDeleteRow() returns error? {
    // Deleting a data row shrinks the table; the row below is pulled up to close the gap.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"], ["Cara", "40"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    check t.deleteRow(1);  // remove Bob
    string[][] rows = check t.getRows();
    test:assertEquals(rows.length(), 2, "Table shrank by one data row");
    test:assertEquals(rows[0][0], "Alice", "Row before the delete preserved");
    test:assertEquals(rows[1][0], "Cara", "Row after the delete pulled up");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableDeleteRowOutOfRange() returns error? {
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = t.deleteRow(5);
    test:assertTrue(result is InvalidTableRangeError, "Out-of-range data-row index must error");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableDeleteLastRowRefused() returns error? {
    // A table must keep at least one data row, so the last one cannot be deleted.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    Table t = check sheet.createTable("T", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    Error? result = t.deleteRow(0);
    test:assertTrue(result is InvalidTableRangeError, "Deleting the only data row must error");
    test:assertEquals(check t.getRowCount(), 1, "The data row is left intact");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableDeleteRowCarriesTotalsRow() returns error? {
    // The totals row and its value ride along with the upward shift when a data row is deleted.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Sales");
    check sheet.putRows([["Region", "Amount"], ["North", "100"], ["South", "250"], ["East", "75"]]);
    Table t = check sheet.createTable("SalesTable", {
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    check setTotalRowNative(t, 1, 425.0);

    check t.deleteRow(0);  // remove North
    test:assertEquals(check t.getRowCount(), 2, "Two data rows remain");
    test:assertTrue(check t.hasTotalRow(), "Totals row survives the delete");
    map<CellValue> totals = check t.getTotalRow();
    test:assertEquals(totals["Amount"], 425, "Totals value carried up by the shift");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableInsertAtShiftsTotalsRow() returns error? {
    // A mid-table insert shifts the totals row down and preserves its value.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Sales");
    check sheet.putRows([["Region", "Amount"], ["North", "100"], ["South", "250"]]);
    Table t = check sheet.createTable("SalesTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    check setTotalRowNative(t, 1, 350.0);

    check t.putRows([["East", "75"]], tableWriteMode = APPEND, insertAt = 1);
    test:assertEquals(check t.getRowCount(), 3, "Data grew to three rows");
    test:assertTrue(check t.hasTotalRow(), "Totals row survives the insert");
    map<CellValue> totals = check t.getTotalRow();
    test:assertEquals(totals["Amount"], 350, "Totals value preserved after the shift");
    string[][] rows = check t.getRows();
    test:assertEquals(rows[1][0], "East", "Inserted at data-row 1");
    check wb.close();
}

// =============================================================================
// createTableFromData column span + empty rejection
// =============================================================================

@test:Config {groups: ["table"]}
function testCreateTableFromDataKeepsAllDeclaredColumns() returns error? {
    // The table spans every declared column even when the first row omits an optional field
    // (the column count is read from the written header, not the first row's present keys).
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    OptionalAgeRow[] data = [{name: "Alice"}, {name: "Bob", age: 30}];
    Table t = check sheet.createTableFromData("T", data);
    test:assertEquals(check t.getColumnCount(), 2, "Table spans both declared columns");
    string[] headers = check t.getHeaders();
    test:assertTrue(headers.indexOf("age") is int,
            "The optional column is part of the table even though the first row omits it");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testCreateTableFromDataEmptyRejected() returns error? {
    // A table needs at least a header row, so empty data cannot form one.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    string[][] empty = [];
    Table|Error result = sheet.createTableFromData("T", empty);
    test:assertTrue(result is InvalidTableRangeError, "Empty data cannot form a table");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testCreateTableDuplicateNameErrors() returns error? {
    // A table name must be unique within the workbook.
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("S");
    check sheet.putRows([["Name", "Age"], ["Alice", "30"]]);
    _ = check sheet.createTable("Dup", {
        firstRowIndex: 0,
        lastRowIndex: 1,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    Table|Error result = sheet.createTable("Dup", {
        firstRowIndex: 3,
        lastRowIndex: 4,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });
    test:assertTrue(result is TableExistsError, "A duplicate table name must raise TableExistsError");
    check wb.close();
}
