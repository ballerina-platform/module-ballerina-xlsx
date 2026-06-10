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
// FAIL-SAFE TEST RECORD TYPES
// =============================================================================

// Record type for fail-safe testing - requires int for age
// Uses @xlsx:Name to match the headers in failsafe_test.xlsx (lowercase)
type FailSafeEmployee record {|
    @Name {value: "name"}
    string name;
    @Name {value: "age"}
    int age;  // Will fail when parsing "invalid" or "abc"
    @Name {value: "department"}
    string department;
|};

// =============================================================================
// TEST DATA CONSTANTS
// =============================================================================

const string FAIL_SAFE_TEST_DIR = "tests/resources/testdata/";
const string FAIL_SAFE_ERROR_LOG = "tests/resources/testdata/failsafe_errors.log";

// =============================================================================
// FAIL-SAFE TESTS
// =============================================================================

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData
}
function testFailFastModeDefault() returns error? {
    // Without fail-safe, parsing should throw on first error
    FailSafeEmployee[]|error result = parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx");

    // Should be a TypeConversionError because "invalid" can't convert to int
    test:assertTrue(result is TypeConversionError, "Should be TypeConversionError");

    if result is TypeConversionError {
        // Verify error details
        ErrorDetails details = result.detail();

        // "Jane" with the "invalid" age is the first error (POI row index 2:
        // index 0 = header, 1 = John, 2 = Jane). "age" is column index 1.
        test:assertEquals(details.rowNumber, 2, "Error row should be Jane's row (index 2)");
        test:assertEquals(details.columnNumber, 1, "Error column should be the age column (index 1)");

        // Error message should mention the conversion failure
        string message = result.message();
        test:assertTrue(message.includes("int") || message.includes("convert"),
            "Error message should mention type conversion issue");
    }
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData
}
function testFailSafeWithConsoleLogs() returns error? {
    // With fail-safe enabled, should continue processing and log errors
    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: true,
            includeSourceDataInConsole: true
        }
    };

    FailSafeEmployee[] employees = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Should have parsed the valid rows (skipped invalid ones)
    // The test file has 4 rows: 2 valid, 2 with invalid ages
    test:assertEquals(employees.length(), 2, "Should have 2 valid employees (2 invalid rows skipped)");

    // Verify the valid rows were parsed correctly
    test:assertEquals(employees[0].name, "John", "First employee should be John");
    test:assertEquals(employees[0].age, 30, "John's age should be 30");
    test:assertEquals(employees[1].name, "Carol", "Second employee should be Carol");
    test:assertEquals(employees[1].age, 28, "Carol's age should be 28");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeWithFileLogging() returns error? {
    // Clean up any existing log file first
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,
                fileWriteOption: OVERWRITE
            }
        }
    };

    FailSafeEmployee[] employees = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Verify parsing succeeded with valid rows
    test:assertEquals(employees.length(), 2, "Should have 2 valid employees");

    // Verify log file was created
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS), "Error log file should exist");

    // Verify log file has content
    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    test:assertTrue(logContent.length() > 0, "Log file should have content");

    // Verify log content contains the metadata JSON keys
    test:assertTrue(logContent.includes("\"message\""), "Log should contain message JSON key");
    test:assertTrue(logContent.includes("\"location\""), "Log should contain location JSON key");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeFileLoggingRawMode() returns error? {
    // Clean up any existing log file first
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW,
                fileWriteOption: OVERWRITE
            }
        }
    };

    // Parse the file - we're testing that errors go to the log file, not the parsed result
    FailSafeEmployee[] _ = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Verify log file was created with RAW content
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS), "Error log file should exist");

    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    // RAW mode should contain the raw row data (JSON array format)
    test:assertTrue(logContent.includes("["), "RAW log should contain JSON array");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData
}
function testFailSafeConsoleLogsDisabled() returns error? {
    // Fail-safe with console logs disabled, no file output
    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false
        }
    };

    FailSafeEmployee[] result = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Should still skip invalid rows even without logging
    test:assertEquals(result.length(), 2, "Should have 2 valid employees");
}

@test:Config {
    groups: ["failsafe"]
}
function testFailSafeWithEmptyConfig() returns error? {
    // Empty fail-safe config should enable fail-safe with defaults
    ParseOptions opts = {
        failSafe: {}
    };

    FailSafeEmployee[] employees = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Should have parsed valid rows
    test:assertEquals(employees.length(), 2, "Should have 2 valid employees with empty failSafe config");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeConsoleAndFileTogether() returns error? {
    // Test that both console and file logging can be enabled simultaneously
    // Clean up any existing log file first
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: true,           // Enable console logging
            includeSourceDataInConsole: true,  // Include row data in console
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,
                fileWriteOption: OVERWRITE
            }
        }
    };

    FailSafeEmployee[] employees = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Verify parsing succeeded with valid rows
    test:assertEquals(employees.length(), 2, "Should have 2 valid employees");

    // Verify log file was created (indicates error handling pipeline works)
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS),
        "Error log file should exist when both console and file logging enabled");

    // Verify log file has error content
    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    test:assertTrue(logContent.length() > 0, "Log file should have content");

    // When RAW_AND_METADATA is used with file logging, it should contain all JSON keys
    test:assertTrue(logContent.includes("\"message\""), "Log should contain message JSON key");
    test:assertTrue(logContent.includes("\"location\""), "Log should contain location JSON key");
    test:assertTrue(logContent.includes("\"offendingRow\""), "RAW_AND_METADATA should include offendingRow key");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeIncludeSourceDataInConsole() returns error? {
    // Test that includeSourceDataInConsole captures row data correctly
    // We verify this indirectly through file logging since we can't capture console output
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: true,
            includeSourceDataInConsole: true,  // This setting also affects error data capture
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,  // Request raw data in file
                fileWriteOption: OVERWRITE
            }
        }
    };

    FailSafeEmployee[] _ = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Verify log file contains offending row data
    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);

    // The offending rows contain "invalid" and "abc" in the age column
    // RAW_AND_METADATA mode should include the raw row data
    test:assertTrue(logContent.includes("offendingRow"),
        "Log should include offendingRow field with source data");

    // The offending row should contain the invalid value
    test:assertTrue(logContent.includes("invalid") || logContent.includes("Jane"),
        "offendingRow should contain data from the failing row");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeMetadataOnlyMode() returns error? {
    // Test METADATA content type - should NOT include raw row data
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: METADATA,  // Metadata only - no raw data
                fileWriteOption: OVERWRITE
            }
        }
    };

    FailSafeEmployee[] _ = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);

    // Verify log file was created
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS),
        "Error log file should exist");

    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);

    // METADATA mode must include all of timestamp, location, and message. Match the
    // quoted JSON keys, not bare substrings, so a value can't pass these by accident.
    test:assertTrue(logContent.includes("\"time\""), "METADATA log should contain time JSON key");
    test:assertTrue(logContent.includes("\"location\""), "METADATA log should contain location JSON key");
    test:assertTrue(logContent.includes("\"message\""), "METADATA log should contain message JSON key");

    // METADATA mode should NOT include offendingRow
    test:assertFalse(logContent.includes("\"offendingRow\""),
        "METADATA mode should NOT include offendingRow field");
}

@test:Config {
    groups: ["failsafe"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testFailSafeFileAppendMode() returns error? {
    // APPEND mode should accumulate entries across parse runs into the same file.
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,
                fileWriteOption: APPEND
            }
        }
    };

    // First run writes the offending-row entries.
    FailSafeEmployee[] _ = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);
    string afterFirst = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    int firstLength = afterFirst.length();
    test:assertTrue(firstLength > 0, "First run should produce log content");

    // Second run with APPEND must add to the existing file, not replace it.
    FailSafeEmployee[] _ = check parseSheet(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx", 0, opts);
    string afterSecond = check io:fileReadString(FAIL_SAFE_ERROR_LOG);

    test:assertTrue(afterSecond.length() > firstLength,
        "APPEND mode should grow the log file across runs");
    test:assertTrue(afterSecond.startsWith(afterFirst),
        "APPEND mode should retain the first run's entries");
}

@test:Config {
    groups: ["failsafe"]
}
function testFailSafeDoesNotSwallowStructuralError() returns error? {
    // Fail-safe only skips row-level errors. A structural error (missing file)
    // must still fail, not be silently swallowed.
    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false
        }
    };

    FailSafeEmployee[]|Error result = parseSheet(FAIL_SAFE_TEST_DIR + "does_not_exist.xlsx", 0, opts);
    test:assertTrue(result is Error,
        "Structural error must fail even with fail-safe enabled");
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

// Setup function to create test data file with some invalid rows
function setupFailSafeTestData() returns error? {
    string testFilePath = FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx";

    // Always recreate the file to ensure consistent headers
    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    // Create test data with mix of valid and invalid rows
    // Columns: name, age, department
    // Row 2: John, 30, Engineering (valid)
    // Row 3: Jane, "invalid", Marketing (invalid - age is not a number)
    // Row 4: Carol, 28, Sales (valid)
    // Row 5: Bob, "abc", HR (invalid - age is not a number)
    string[][] testData = [
        ["name", "age", "department"],  // Header row
        ["John", "30", "Engineering"],   // Valid
        ["Jane", "invalid", "Marketing"], // Invalid age
        ["Carol", "28", "Sales"],        // Valid
        ["Bob", "abc", "HR"]              // Invalid age
    ];

    check writeSheet(testData, testFilePath);
}

// Setup function to create a TABLE fixture with invalid rows
function setupFailSafeTableTestData() returns error? {
    string testFilePath = FAIL_SAFE_TEST_DIR + "failsafe_table.xlsx";

    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    string[][] testData = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],    // Valid
        ["Jane", "invalid", "Marketing"], // Invalid age
        ["Carol", "28", "Sales"],         // Valid
        ["Bob", "abc", "HR"]              // Invalid age
    ];
    check sheet.putRows(testData);
    _ = check sheet.createTable("FailSafeTable", {
        firstRowIndex: 0,
        lastRowIndex: 4,
        firstColumnIndex: 0,
        lastColumnIndex: 2
    });
    check wb.saveAs(testFilePath);
    check wb.close();
}

// Cleanup function to remove error log file
function cleanupFailSafeLogFile() returns error? {
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }
}

// =============================================================================
// SHEET API FAIL-SAFE TESTS
// =============================================================================
// Tests for fail-safe support via the Workbook/Sheet API

@test:Config {
    groups: ["failsafe", "workbook"],
    before: setupFailSafeTestData
}
function testSheetGetRowsWithFailSafe() returns error? {
    // Test that Sheet.getRows() supports fail-safe error handling
    Workbook wb = check fromFile(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx");
    Sheet sheet = check wb.getSheet(0);

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false
        }
    };

    // With fail-safe enabled via Sheet API, invalid rows should be skipped
    FailSafeEmployee[] employees = check sheet.getRows(opts);

    // Should have 2 valid employees (2 invalid rows skipped)
    test:assertEquals(employees.length(), 2,
        "Sheet.getRows with failSafe should skip invalid rows");
    test:assertEquals(employees[0].name, "John", "First employee should be John");
    test:assertEquals(employees[1].name, "Carol", "Second employee should be Carol");

    check wb.close();
}

@test:Config {
    groups: ["failsafe", "workbook"],
    before: setupFailSafeTestData,
    after: cleanupFailSafeLogFile
}
function testSheetGetRowsWithFailSafeFileLogging() returns error? {
    // Test that Sheet.getRows() with fail-safe supports file logging
    // Clean up any existing log file first
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    Workbook wb = check fromFile(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx");
    Sheet sheet = check wb.getSheet(0);

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,
                fileWriteOption: OVERWRITE
            }
        }
    };

    FailSafeEmployee[] employees = check sheet.getRows(opts);

    test:assertEquals(employees.length(), 2, "Should have 2 valid employees");

    // Verify log file was created
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS),
        "Error log should be created for Sheet.getRows with failSafe");

    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    test:assertTrue(logContent.includes("\"message\""), "Log should contain the message JSON key");

    check wb.close();
}

@test:Config {
    groups: ["failsafe", "workbook"],
    before: setupFailSafeTestData
}
function testSheetGetRowsWithoutFailSafeFails() returns error? {
    // Test that Sheet.getRows() without fail-safe fails on first error
    Workbook wb = check fromFile(FAIL_SAFE_TEST_DIR + "failsafe_test.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // Without fail-safe, parsing should throw on first error
    FailSafeEmployee[]|Error result = sheet.getRows();

    test:assertTrue(result is TypeConversionError,
        "Sheet.getRows without failSafe should return TypeConversionError");

    check wb.close();
}

// =============================================================================
// TABLE API FAIL-SAFE TESTS
// =============================================================================
// Fail-safe must work on table reads too, not only parseSheet / Sheet.getRows.

@test:Config {
    groups: ["failsafe", "table"],
    before: setupFailSafeTableTestData,
    after: cleanupFailSafeLogFile
}
function testTableGetRowsWithFailSafe() returns error? {
    if check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS) {
        check file:remove(FAIL_SAFE_ERROR_LOG);
    }

    Workbook wb = check fromFile(FAIL_SAFE_TEST_DIR + "failsafe_table.xlsx");
    Table tbl = check wb.getTable("FailSafeTable");

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: FAIL_SAFE_ERROR_LOG,
                contentType: RAW_AND_METADATA,
                fileWriteOption: OVERWRITE
            }
        }
    };

    // With fail-safe, the two non-numeric age rows must be skipped, not abort the read.
    FailSafeEmployee[] employees = check tbl.getRows(opts);

    test:assertEquals(employees.length(), 2,
        "Table.getRows with failSafe should skip the 2 invalid rows");
    test:assertEquals(employees[0].name, "John", "First surviving row should be John");
    test:assertEquals(employees[1].name, "Carol", "Second surviving row should be Carol");

    // The skipped rows must be logged.
    test:assertTrue(check file:test(FAIL_SAFE_ERROR_LOG, file:EXISTS),
        "Error log should be created for Table.getRows with failSafe");
    string logContent = check io:fileReadString(FAIL_SAFE_ERROR_LOG);
    test:assertTrue(logContent.includes("\"message\""), "Log should contain the message JSON key");

    check wb.close();
}

// =============================================================================
// BLANK CELL HANDLING WITH FAIL-SAFE TESTS
// =============================================================================

@test:Config {
    groups: ["failsafe"]
}
function testBlankCellRequiredFieldWithFailSafe() returns error? {
    // With fail-safe, blank cell for required field should skip row, not fail
    string testFile = FAIL_SAFE_TEST_DIR + "blank_failsafe_test.xlsx";
    string[][] data = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],
        ["Jane", "", "Marketing"],    // Blank age - will be skipped
        ["Carol", "28", "Sales"]
    ];
    check writeSheet(data, testFile);

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false
        }
    };

    FailSafeEmployee[] result = check parseSheet(testFile, 0, opts);

    // Jane should be skipped (blank required field), John and Carol should be parsed
    test:assertEquals(result.length(), 2, "Should have 2 records (blank row skipped)");
    test:assertEquals(result[0].name, "John", "First should be John");
    test:assertEquals(result[1].name, "Carol", "Second should be Carol");

    check file:remove(testFile);
}

@test:Config {
    groups: ["failsafe", "workbook"]
}
function testSheetBlankCellRequiredFieldWithFailSafe() returns error? {
    // Test blank cell handling with fail-safe via Sheet API
    string testFile = FAIL_SAFE_TEST_DIR + "sheet_blank_failsafe_test.xlsx";
    string[][] data = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],
        ["Jane", "", "Marketing"],    // Blank age - will be skipped
        ["Carol", "28", "Sales"]
    ];
    check writeSheet(data, testFile);

    Workbook wb = check fromFile(testFile);
    Sheet sheet = check wb.getSheet(0);

    ParseOptions opts = {
        failSafe: {
            enableConsoleLogs: false
        }
    };

    FailSafeEmployee[] result = check sheet.getRows(opts);

    // Jane should be skipped (blank required field)
    test:assertEquals(result.length(), 2, "Sheet.getRows should skip blank required field rows");
    test:assertEquals(result[0].name, "John", "First should be John");
    test:assertEquals(result[1].name, "Carol", "Second should be Carol");

    check wb.close();
    check file:remove(testFile);
}

@test:AfterSuite
function cleanupFailSafeTestData() returns error? {
    string[] filesToRemove = [
        "failsafe_test.xlsx",
        "failsafe_table.xlsx",
        "blank_failsafe_test.xlsx",
        "sheet_blank_failsafe_test.xlsx",
        "failsafe_errors.log"
    ];

    foreach string fileName in filesToRemove {
        string filePath = FAIL_SAFE_TEST_DIR + fileName;
        if check file:test(filePath, file:EXISTS) {
            check file:remove(filePath);
        }
    }
}
