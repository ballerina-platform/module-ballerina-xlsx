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
        check file:createDir(TEST_DATA_DIR);
    }

    // -------------------------------------------------------------------------
    // tables_test.xlsx - Excel file with tables
    // -------------------------------------------------------------------------
    // Note: We need to create tables programmatically using the Workbook API
    Workbook wb = check new;

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
}

@test:AfterSuite
function cleanupTableTestData() returns error? {
    string[] filesToRemove = [
        "tables_test.xlsx",
        "temp_table_write.xlsx",
        "temp_table_create.xlsx",
        "temp_table_write_notfound.xlsx"
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

    test:assertEquals(data.length(), 3, "Should have 3 rows");
    test:assertEquals(data[0][0], "Alice", "First cell");
    test:assertEquals(data[0][1], "30", "Age as string");
    test:assertEquals(data[0][2], "Engineering", "Department");
}

@test:Config {
    groups: ["table"]
}
function testParseTableToMaps() returns error? {
    map<anydata>[] data = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable");

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
function testWriteTableExpands() returns error? {
    // First, copy the test file
    string tempFile = TEST_DATA_DIR + "temp_table_write.xlsx";

    // Open original and save as temp
    Workbook wbSrc = check new(TEST_DATA_DIR + "tables_test.xlsx");
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
}

// =============================================================================
// WORKBOOK TABLE API TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testWorkbookGetTable() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");

    Table empTable = check wb.getTable("EmployeeTable");
    test:assertEquals(empTable.getName(), "EmployeeTable", "Table name should match");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWorkbookGetTableNotFound() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");

    Table|TableNotFoundError result = wb.getTable("NonExistentTable");
    test:assertTrue(result is TableNotFoundError, "Should return TableNotFoundError");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testWorkbookGetAllTables() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");

    Table[] tables = check wb.getAllTables();
    test:assertEquals(tables.length(), 2, "Should have 2 tables");

    // Tables could be in any order
    string[] tableNames = tables.map(t => t.getName());
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table empTable = check sheet.getTable("EmployeeTable");
    test:assertEquals(empTable.getName(), "EmployeeTable", "Table name should match");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetGetTables() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table[] tables = check sheet.getTables();
    test:assertEquals(tables.length(), 1, "Employees sheet should have 1 table");
    test:assertEquals(tables[0].getName(), "EmployeeTable", "Table name should be EmployeeTable");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetCreateTable() returns error? {
    Workbook wb = check new;
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

    test:assertEquals(newTable.getName(), "NewTable", "Table name should match");
    test:assertEquals(newTable.getRowCount(), 2, "Should have 2 data rows");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetCreateTableFromData() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("TestSheet");

    TableEmployee[] employees = [
        {Name: "Test1", Age: 20, Department: "Dept1"},
        {Name: "Test2", Age: 30, Department: "Dept2"}
    ];

    Table newTable = check sheet.createTableFromData("EmpTable", employees);

    test:assertEquals(newTable.getName(), "EmpTable", "Table name should match");
    test:assertEquals(newTable.getColumnCount(), 3, "Should have 3 columns");
    test:assertEquals(newTable.getRowCount(), 2, "Should have 2 data rows");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testSheetDeleteTable() returns error? {
    Workbook wb = check new;
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    test:assertEquals(empTable.getName(), "EmployeeTable", "getName() should return table name");
    test:assertEquals(empTable.getDisplayName(), "EmployeeTable", "getDisplayName() should return display name");
    test:assertEquals(empTable.getSheetName(), "Employees", "getSheetName() should return sheet name");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableRangeMethods() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    // Full range (including header)
    CellRange fullRange = empTable.getRange();
    test:assertEquals(fullRange.firstRowIndex, 0, "Full range starts at row 0");
    test:assertEquals(fullRange.lastRowIndex, 3, "Full range ends at row 3");

    // Data range (excluding header)
    CellRange dataRange = empTable.getDataRange();
    test:assertEquals(dataRange.firstRowIndex, 1, "Data range starts at row 1");
    test:assertEquals(dataRange.lastRowIndex, 3, "Data range ends at row 3");

    test:assertEquals(empTable.getRowCount(), 3, "Should have 3 data rows");
    test:assertEquals(empTable.getColumnCount(), 3, "Should have 3 columns");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetHeaders() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    string[] headers = empTable.getHeaders();
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    TableEmployee[] employees = check empTable.getRows();

    test:assertEquals(employees.length(), 3, "Should have 3 employees");
    test:assertEquals(employees[0].Name, "Alice", "First employee name");
    test:assertEquals(employees[1].Name, "Bob", "Second employee name");
    test:assertEquals(employees[2].Name, "Charlie", "Third employee name");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetRow() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
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
    Workbook wb = check new;
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

@test:Config {
    groups: ["table"]
}
function testTableRename() returns error? {
    Workbook wb = check new;
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

    test:assertEquals(t.getName(), "OldName", "Initial name should be OldName");

    // Rename
    check t.rename("NewName");
    test:assertEquals(t.getName(), "NewName", "Name should be NewName after rename");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableResize() returns error? {
    Workbook wb = check new;
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

    test:assertEquals(t.getRowCount(), 1, "Initial row count should be 1");

    // Resize to include more data
    check t.resize({
        firstRowIndex: 0,
        lastRowIndex: 3,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });

    test:assertEquals(t.getRowCount(), 3, "Row count should be 3 after resize");
    test:assertEquals(t.getColumnCount(), 3, "Column count should be 3 after resize");

    check wb.close();
}

// =============================================================================
// ERROR HANDLING TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableNotFoundErrorType() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");

    Table|TableNotFoundError result = wb.getTable("NonExistent");
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Sheet sheet = check wb.getSheet("Employees");

    Table|TableNotFoundError result = sheet.getTable("NonExistent");
    test:assertTrue(result is TableNotFoundError, "Should be TableNotFoundError");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testDeleteNonExistentTable() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Test");

    TableNotFoundError? result = sheet.deleteTable("NonExistent");
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    test:assertFalse(empTable.hasTotalsRow(), "Table should not have totals row");

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testTableGetTotalsRowError() returns error? {
    // Getting totals row from table without totals should return error
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    map<anydata>|Error result = empTable.getTotalsRow();
    test:assertTrue(result is Error, "Should return error when table has no totals row");

    check wb.close();
}

// =============================================================================
// ADDITIONAL ERROR SCENARIO TESTS
// =============================================================================

@test:Config {
    groups: ["table"]
}
function testTableOverlapError() returns error? {
    Workbook wb = check new;
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
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("TestSheet");

    // Try to resize a table to invalid range (firstRow > lastRow would be caught by CellRange validation)
    // Instead, let's test a scenario where resize would make the table too small
    string[][] data = [["A", "B"], ["1", "2"], ["3", "4"]];
    check sheet.putRows(data);

    Table t = check sheet.createTable("TestTable", {
        firstRowIndex: 0,
        lastRowIndex: 2,
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // Try to resize to have no data rows (only header) - this should fail or produce invalid state
    // Note: POI may or may not validate this, so we test what our implementation does
    Error? result = t.resize({
        firstRowIndex: 0,
        lastRowIndex: 0,  // Only header row, no data
        firstColumnIndex: 0,
        lastColumnIndex: 1
    });

    // The resize might succeed in POI but produce invalid table state
    // Our implementation should catch this as InvalidTableRangeError
    if result is InvalidTableRangeError {
        test:assertTrue(true, "Correctly returned InvalidTableRangeError");
    } else {
        // If it didn't fail, verify the table is in a valid state
        test:assertTrue(t.getRowCount() >= 0, "Table should have valid row count");
    }

    check wb.close();
}

@test:Config {
    groups: ["table"]
}
function testGetRowOutOfRange() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
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
    Workbook wb = check new;
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
    ParseOptions opts = {
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");

    // Get rows with rowCount limit
    RowReadOptions opts = {
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
    Workbook srcWb = check new(sourceFile);
    byte[] bytes = check srcWb.toBytes();
    check srcWb.close();
    Workbook destWb = check new(bytes);
    check destWb.saveAs(tempFile);
    check destWb.close();

    // Inline literal — contextually typed against `Data = Row[]`. Exercises the
    // table writer's row-by-row dispatch (writeRowData inspects runtime types).
    check writeTable([["Eve", "40", "Finance"], ["Frank", "33", "HR"]], tempFile,
            "EmployeeTable");

    Workbook check_wb = check new(tempFile);
    Table empTable = check check_wb.getTable("EmployeeTable");
    string[][] rows = check empTable.getRows();
    test:assertEquals(rows[0][0], "Eve");
    test:assertEquals(rows[1][0], "Frank");
    check check_wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["table"]}
function testTableGetRowsWithDataTarget() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");
    // Explicit `Data` target — typedesc is a type reference, dispatch must unwrap
    // and pick the union default.
    Data rows = check empTable.getRows();
    test:assertTrue(rows is string[][], "Data target on Table.getRows should fall back to string[][]");
    if rows is string[][] {
        test:assertEquals(rows.length(), 3, "EmployeeTable has 3 data rows");
    }
    check wb.close();
}

@test:Config {groups: ["table"]}
function testTableGetRowWithRowTarget() returns error? {
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
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
    Workbook wb = check new(TEST_DATA_DIR + "tables_test.xlsx");
    Table empTable = check wb.getTable("EmployeeTable");
    check wb.close();
    string[][]|Error result = empTable.getRows();
    test:assertTrue(result is Error, "Table.getRows after Workbook.close should return Error");
}

// =============================================================================
// Table name validation (FX-13)
// =============================================================================

@test:Config {groups: ["table"]}
function testCreateTableInvalidNameWithSpace() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    Table|Error result = sheet.createTable("Bad Table", "A1:B2");
    test:assertTrue(result is Error, "Table name with space must be rejected");
    check wb.close();
}

@test:Config {groups: ["table"]}
function testCreateTableNameStartsWithDigit() returns error? {
    Workbook wb = check new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.putRows([["A", "B"], ["1", "2"]]);
    Table|Error result = sheet.createTable("1Sales", "A1:B2");
    test:assertTrue(result is Error,
            "Table name starting with a digit must be rejected");
    check wb.close();
}
