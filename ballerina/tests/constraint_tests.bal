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
// These tests verify that the enableConstraintValidation option is properly
// handled. When the constraint module is not available (as in this test env),
// validation is gracefully skipped.

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
    // Parse with constraint validation enabled (default)
    // Should work even if constraint module is not available
    Employee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_test.xlsx");

    // Should parse successfully
    test:assertEquals(employees.length(), 3, "Should parse all 3 rows");
    test:assertEquals(employees[0].name, "John");
    test:assertEquals(employees[0].age, 30);
}

@test:Config {
    groups: ["constraint"],
    before: setupConstraintTestData
}
function testConstraintValidationExplicitlyEnabled() returns error? {
    // Explicitly enable constraint validation
    ParseOptions opts = {
        enableConstraintValidation: true
    };

    Employee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

    // Should parse successfully
    test:assertEquals(employees.length(), 3, "Should parse all 3 rows");
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
    before: setupConstraintTestData
}
function testConstraintValidationWithFailSafe() returns error? {
    // Test that constraint validation works with fail-safe mode
    ParseOptions opts = {
        enableConstraintValidation: true,
        failSafe: {
            enableConsoleLogs: false
        }
    };

    Employee[] employees = check parseSheet(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

    // Should parse successfully
    test:assertEquals(employees.length(), 3, "Should parse all 3 rows");
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

@test:AfterSuite
function cleanupConstraintTestData() returns error? {
    string[] filesToRemove = [
        "constraint_test.xlsx",
        "constraint_violation_test.xlsx"
    ];

    foreach string fileName in filesToRemove {
        string filePath = CONSTRAINT_TEST_DIR + fileName;
        if check file:test(filePath, file:EXISTS) {
            check file:remove(filePath);
        }
    }
}
