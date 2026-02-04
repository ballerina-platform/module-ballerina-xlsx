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
    Employee[] employees = check parse(CONSTRAINT_TEST_DIR + "constraint_test.xlsx");

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

    Employee[] employees = check parse(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

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

    Employee[] employees = check parse(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

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

    Employee[] employees = check parse(CONSTRAINT_TEST_DIR + "constraint_test.xlsx", 0, opts);

    // Should parse successfully
    test:assertEquals(employees.length(), 3, "Should parse all 3 rows");
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

    check write(testData, testFilePath);
}
