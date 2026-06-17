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

import ballerina/constraint;
import ballerina/file;
import ballerina/test;

// =============================================================================
// CONSTRAINT VALIDATION OPTION TESTS
// =============================================================================
// These tests verify the enableConstraintValidation option. Validation runs by
// default against any @constraint annotations: out-of-range rows are rejected
// (or skipped under fail-safe), and disabling validation lets them through.

// =============================================================================
// TEST CONSTANTS
// =============================================================================

const string CONSTRAINT_TEST_DIR = "tests/resources/testdata/";

// =============================================================================
// CONSTRAINT VALIDATION OPTION TESTS
// =============================================================================

@test:Config {
    groups: ["constraint"],
    before: setupConstraintTestData
}
function testConstraintValidationEnabledDefault() returns error? {
    // Validation is enabled by default. constraint_test.xlsx holds only in-range ages
    // (30, 25, 45), so the default-on validator actually runs and all three rows pass.
    ConstrainedEmployee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_test.xlsx");

    test:assertEquals(employees.length(), 3, "All 3 in-range rows should pass default validation");
    test:assertEquals(employees[0].name, "John");
    test:assertEquals(employees[0].age, 30);
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintViolationTestData
}
function testConstraintValidationExplicitlyEnabled() returns error? {
    // Explicitly enable constraint validation against the violation fixture.
    // Because validation actually runs, the out-of-range rows must make parsing
    // fail rather than slip through — a no-op validator would pass this.
    ParseOptions opts = {
        enableConstraintValidation: true
    };

    ConstrainedEmployee[]|Error result =
            parseSheet(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx", 0, opts);

    test:assertTrue(result is ConstraintValidationError,
            "Enabled validation must reject the out-of-range rows");
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintTestData
}
function testConstraintValidationDisabled() returns error? {
    // Explicitly disable constraint validation
    ParseOptions opts = {
        enableConstraintValidation: false
    };

    Employee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

    // Should parse successfully
    test:assertEquals(employees.length(), 3, "Should parse all 3 rows");
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintViolationTestData
}
function testConstraintValidationWithFailSafe() returns error? {
    // Constraint validation combined with fail-safe must SKIP the violating rows
    // (age -5 and 150) rather than fail or pass everything through.
    ParseOptions opts = {
        enableConstraintValidation: true,
        failSafe: {
            enableConsoleLogs: false
        }
    };

    ConstrainedEmployee[] employees =
            check parseSheet(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx", 0, opts);

    test:assertEquals(employees.length(), 2, "Only the two in-range rows must survive");
    test:assertEquals(employees[0].name, "John", "First surviving row is John (age 30)");
    test:assertEquals(employees[1].name, "Bob", "Second surviving row is Bob (age 45)");
}

// =============================================================================
// ACTUAL CONSTRAINT VIOLATION TESTS
// =============================================================================
// These tests verify that constraint violations are properly detected and reported.

// Type with age constraint (must be between 18 and 120)
type ConstrainedEmployee record {|
    string name;
    @constraint:Int {minValue: 18, maxValue: 120}
    int age;
    string department;
|};

@test:Config {
    groups: ["constraint"],
    before: setupConstraintViolationTestData
}
function testConstraintViolationReturnsError() returns error? {
    // Parse data with constraint violation (age = -5)
    // Should return ConstraintValidationError
    ConstrainedEmployee[]|Error result = parseSheet(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx");

    test:assertTrue(result is ConstraintValidationError,
        "Should return ConstraintValidationError for invalid age");

    if result is ConstraintValidationError {
        // The error must pinpoint the offending row, not just signal a generic failure.
        test:assertTrue(result.detail().rowNumber is int,
                "ConstraintValidationError should carry the violating row number");
    }
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintViolationTestData
}
function testConstraintViolationWithFailSafeSkipsInvalidRows() returns error? {
    // With failSafe enabled, invalid rows should be skipped
    ParseOptions opts = {
        enableConstraintValidation: true,
        failSafe: {
            enableConsoleLogs: false
        }
    };

    ConstrainedEmployee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx", 0, opts);

    // Should only return valid rows (rows with age between 18 and 120)
    test:assertEquals(employees.length(), 2, "Should skip rows with invalid age");
    test:assertEquals(employees[0].name, "John");  // age 30 - valid
    test:assertEquals(employees[1].name, "Bob");   // age 45 - valid
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintViolationTestData
}
function testConstraintViolationDisabledAllowsInvalidData() returns error? {
    // With constraint validation disabled, invalid data should pass through
    ParseOptions opts = {
        enableConstraintValidation: false
    };

    ConstrainedEmployee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx", 0, opts);

    // All rows should be returned including invalid ones
    test:assertEquals(employees.length(), 4, "Should return all rows when validation disabled");
}

// =============================================================================
// CONSTRAINT VALIDATION ACROSS TABLE AND SINGLE-ROW READ PATHS
// =============================================================================
// Constraint validation must apply uniformly — not only on parseSheet/Sheet.getRows,
// but also on the table read paths and Sheet.getRow.

@test:Config {
    groups: ["constraint", "table"],
    before: setupConstraintTableTestData
}
function testParseTableConstraintValidationReturnsError() returns error? {
    // parseTable must honour enableConstraintValidation just like parseSheet.
    // The table holds out-of-range ages (-5, 150), so enabled validation must reject them.
    TableParseOptions opts = {
        enableConstraintValidation: true
    };

    ConstrainedEmployee[]|Error result =
            parseTable(CONSTRAINT_TEST_DIR + "constraint_table.xlsx", "ConstraintTable", opts);

    test:assertTrue(result is ConstraintValidationError,
            "parseTable with constraint validation must reject the out-of-range rows");
}

@test:Config {
    groups: ["constraint", "table"],
    before: setupConstraintTableTestData
}
function testTableGetRowsConstraintValidationReturnsError() returns error? {
    // The object-API table read path (Table.getRows) must validate constraints too.
    Workbook wb = check fromFile(CONSTRAINT_TEST_DIR + "constraint_table.xlsx");
    Table tbl = check wb.getTable("ConstraintTable");

    TableParseOptions opts = {
        enableConstraintValidation: true
    };

    ConstrainedEmployee[]|Error result = tbl.getRows(opts);

    test:assertTrue(result is ConstraintValidationError,
            "Table.getRows with constraint validation must reject the out-of-range rows");

    check wb.close();
}

@test:Config {
    groups: ["constraint", "sheet"],
    before: setupConstraintViolationTestData
}
function testSheetGetRowConstraintValidationReturnsError() returns error? {
    // Sheet.getRow must validate constraints. Data-row index 1 is Jane (age -5),
    // which violates the 18..120 bound.
    Workbook wb = check fromFile(CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx");
    Sheet sheet = check wb.getSheet(0);

    RowParseOptions opts = {
        enableConstraintValidation: true
    };

    ConstrainedEmployee|Error result = sheet.getRow(1, opts);

    test:assertTrue(result is ConstraintValidationError,
            "Sheet.getRow with constraint validation must reject the out-of-range row");

    check wb.close();
}

// Setup function for constraint violation test data
function setupConstraintViolationTestData() returns error? {
    string testFilePath = CONSTRAINT_TEST_DIR + "constraint_violation_test.xlsx";

    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    // Create test data with constraint violations
    // age must be between 18 and 120
    string[][] testData = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],   // Valid
        ["Jane", "-5", "Marketing"],     // Invalid: age < 18
        ["Bob", "45", "Sales"],          // Valid
        ["Alice", "150", "HR"]           // Invalid: age > 120
    ];

    check writeSheet(testData, testFilePath);
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

// Setup function to create test data
function setupConstraintTestData() returns error? {
    string testFilePath = CONSTRAINT_TEST_DIR + "constraint_test.xlsx";

    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    // Create test data
    string[][] testData = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],
        ["Jane", "25", "Marketing"],
        ["Bob", "45", "Sales"]
    ];

    check writeSheet(testData, testFilePath);
}

// Setup function for a table fixture carrying constraint-violating rows
function setupConstraintTableTestData() returns error? {
    string testFilePath = CONSTRAINT_TEST_DIR + "constraint_table.xlsx";

    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    string[][] testData = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],   // Valid
        ["Jane", "-5", "Marketing"],     // Invalid: age < 18
        ["Bob", "45", "Sales"],          // Valid
        ["Alice", "150", "HR"]           // Invalid: age > 120
    ];
    check sheet.putRows(testData);
    _ = check sheet.createTable("ConstraintTable", {
        firstRowIndex: 0,
        lastRowIndex: 4,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });
    check wb.saveAs(testFilePath);
    check wb.close();
}

@test:AfterSuite
function cleanupConstraintTestData() returns error? {
    string[] filesToRemove = [
        "constraint_test.xlsx",
        "constraint_violation_test.xlsx",
        "constraint_table.xlsx"
    ];

    foreach string fileName in filesToRemove {
        string filePath = CONSTRAINT_TEST_DIR + fileName;
        if check file:test(filePath, file:EXISTS) {
            check file:remove(filePath);
        }
    }
}
