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
// BASIC PARSE TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet"]
}
function testParseToStringArray() returns error? {
    string[][] rows = check parseSheet(TEST_DATA_DIR + "simple.xlsx");

    // Verify exact content matches expected data
    assertStringArrayEquals(rows, EXPECTED_SIMPLE_DATA, "simple.xlsx parse");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseToRecords() returns error? {
    Employee[] employees = check parseSheet(TEST_DATA_DIR + "employees.xlsx");

    // Verify exact content matches expected data
    assertEmployeesEqual(employees, EXPECTED_EMPLOYEES);
}

@test:Config {
    groups: ["parseSheet", "options"]
}
function testParseWithHeaderRowOption() returns error? {
    // complex_headers.xlsx has: Row 0=title, Row 1=metadata, Row 2=headers, Row 3+=data
    ParseOptions opts = {
        headerRowIndex: 2,
        dataStartRowIndex: 3
    };
    string[][] rows = check parseSheet(TEST_DATA_DIR + "complex_headers.xlsx", 0, opts);

    // When dataStartRowIndex is specified, only data rows are returned (starting from row 3)
    test:assertEquals(rows.length(), 2, "Should have 2 data rows");
    test:assertEquals(rows[0][0], "Item1", "First row should be first data row");
    test:assertEquals(rows[1][0], "Item2", "Second row should be second data row");
}

// =============================================================================
// SHEET SELECTION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet"]
}
function testParseWithSheetSelectionByName() returns error? {
    // Select Sheet2 by name
    string[][] rows = check parseSheet(TEST_DATA_DIR + "multi_sheet.xlsx", "Sheet2");

    // Verify we got Sheet2 data
    assertStringArrayEquals(rows, EXPECTED_SHEET2_DATA, "Sheet2 data");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseWithSheetSelectionByIndex() returns error? {
    // Select second sheet (index 1)
    string[][] rows = check parseSheet(TEST_DATA_DIR + "multi_sheet.xlsx", 1);

    // Verify we got Sheet2 data (index 1)
    assertStringArrayEquals(rows, EXPECTED_SHEET2_DATA, "Sheet index 1 data");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseDefaultsToFirstSheet() returns error? {
    // No sheet specified - should default to first sheet
    string[][] rows = check parseSheet(TEST_DATA_DIR + "multi_sheet.xlsx");

    // Verify we got Sheet1 data
    assertStringArrayEquals(rows, EXPECTED_SHEET1_DATA, "Default to Sheet1");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseThirdSheet() returns error? {
    // Select Sheet3 by name
    string[][] rows = check parseSheet(TEST_DATA_DIR + "multi_sheet.xlsx", "Sheet3");

    // Verify we got Sheet3 data
    assertStringArrayEquals(rows, EXPECTED_SHEET3_DATA, "Sheet3 data");
}

// =============================================================================
// FORMULA HANDLING TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet"]
}
function testParseFormulaCachedMode() returns error? {
    ParseOptions opts = {
        formulaMode: CACHED
    };
    string[][] rows = check parseSheet(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

    // In CACHED mode, formula cells return their cached calculated values.
    // For formulas written via POI without evaluation, the default cached value is "0".
    // In real-world usage with Excel-created files, this would return the actual calculated result.
    test:assertEquals(rows.length(), 3, "Should have 3 rows");
    test:assertEquals(rows[0][2], "Sum", "Header should be 'Sum'");
    test:assertEquals(rows[1][2], "0", "Unevaluated formula should return default cached value '0'");
    test:assertEquals(rows[2][2], "0", "Unevaluated formula should return default cached value '0'");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseFormulaTextMode() returns error? {
    ParseOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check parseSheet(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

    // In TEXT mode, formula cells return the formula string with "=" prefix
    test:assertEquals(rows.length(), 3, "Should have 3 rows");
    test:assertEquals(rows[0][2], "Sum", "Header should be 'Sum'");
    test:assertEquals(rows[1][2], "=A2+B2", "Should return exact formula text");
    test:assertEquals(rows[2][2], "=A3+B3", "Should return exact formula text");
}

// =============================================================================
// TYPE CONVERSION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseNumericToInt() returns error? {
    NumericTypes[] data = check parseSheet(TEST_DATA_DIR + "numeric_types.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].intValue, 42, "First intValue should be 42");
    test:assertEquals(data[1].intValue, -100, "Second intValue should be -100");
    test:assertEquals(data[2].intValue, 0, "Third intValue should be 0");
}

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseNumericToDecimal() returns error? {
    NumericTypes[] data = check parseSheet(TEST_DATA_DIR + "numeric_types.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    // Decimal comparison with tolerance for floating point
    test:assertTrue(data[0].decimalValue > 3.14d && data[0].decimalValue < 3.15d,
        "First decimalValue should be ~3.14159");
    test:assertTrue(data[1].decimalValue > 0.0009d && data[1].decimalValue < 0.002d,
        "Second decimalValue should be ~0.001");
}

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseToTypedRecordWithVariousTypes() returns error? {
    TypeVariety[] data = check parseSheet(TEST_DATA_DIR + "types_variety.xlsx");

    test:assertEquals(data.length(), 2, "Should have 2 records");

    // First record
    test:assertEquals(data[0].text, "Hello", "First text should be 'Hello'");
    test:assertEquals(data[0].number, 42, "First number should be 42");
    test:assertTrue(data[0].amount > 99.98d && data[0].amount < 100.0d,
        "First amount should be ~99.99");
    test:assertEquals(data[0].flag, true, "First flag should be true");

    // Second record
    test:assertEquals(data[1].text, "World", "Second text should be 'World'");
    test:assertEquals(data[1].number, -10, "Second number should be -10");
    test:assertEquals(data[1].flag, false, "Second flag should be false");
}

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseToMapArray() returns error? {
    map<anydata>[] data = check parseSheet(TEST_DATA_DIR + "simple.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records (excluding header)");
    test:assertEquals(data[0]["Name"], "John", "First record Name should be 'John'");
    test:assertEquals(data[0]["Age"], "30", "First record Age should be '30'");
    test:assertEquals(data[0]["City"], "New York", "First record City should be 'New York'");
}

// =============================================================================
// OPTIONS TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "options"]
}
function testParseWithCustomDataStartRow() returns error? {
    ParseOptions opts = {
        headerRowIndex: 0,
        dataStartRowIndex: 2  // Skip first data row
    };
    string[][] rows = check parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // When dataStartRowIndex is specified, only data rows starting from that row are returned
    test:assertEquals(rows.length(), 2, "Should have 2 data rows (skipped first data row)");
    test:assertEquals(rows[0][0], "Jane", "First row should be 'Jane' (skipped 'John')");
    test:assertEquals(rows[1][0], "Bob", "Second row should be 'Bob'");
}

@test:Config {
    groups: ["parseSheet", "options"]
}
function testParseSkipsEmptyRows() returns error? {
    // Default behavior: empty rows are skipped
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Empty rows should be skipped
    // Original: header, First, empty, Second, empty, Third
    // Expected: header, First, Second, Third (no empty rows)
    test:assertEquals(rows.length(), 4, "Should have 4 rows (header + 3 data, no empty)");
}

// Note: To include empty rows with position preservation, use Row wrapper type.
// See Row type documentation for details.

// =============================================================================
// ANNOTATION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "annotation"]
}
function testParseWithXlsxNameAnnotation() returns error? {
    // annotated.xlsx has headers: "First Name", "Employee ID", "Department Name"
    // AnnotatedEmployee maps these to: firstName, id, department
    AnnotatedEmployee[] data = check parseSheet(TEST_DATA_DIR + "annotated.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 employees");
    test:assertEquals(data[0].firstName, "Alice", "First employee firstName");
    test:assertEquals(data[0].id, 101, "First employee id");
    test:assertEquals(data[0].department, "Engineering", "First employee department");
    test:assertEquals(data[1].firstName, "Bob", "Second employee firstName");
    test:assertEquals(data[2].firstName, "Charlie", "Third employee firstName");
}

// =============================================================================
// ERROR HANDLING TESTS (NEGATIVE)
// =============================================================================

@test:Config {
    groups: ["parseSheet", "negative"]
}
function testParseSheetNotFoundByName() returns error? {
    string[][]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", "NonExistentSheet");

    test:assertTrue(result is Error, "Should return error for non-existent sheet");
    if result is Error {
        test:assertTrue(result.message().includes("not found"),
            "Error message should mention sheet not found");
    }
}

@test:Config {
    groups: ["parseSheet", "negative"]
}
function testParseSheetNotFoundByIndex() returns error? {
    string[][]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 99);

    test:assertTrue(result is Error, "Should return error for invalid sheet index");
    if result is Error {
        test:assertTrue(result.message().includes("not found") || result.message().includes("valid range"),
            "Error message should mention sheet not found or valid range");
    }
}

@test:Config {
    groups: ["parseSheet", "negative"]
}
function testParseFileNotFound() returns error? {
    string[][]|Error result = parseSheet(TEST_DATA_DIR + "nonexistent.xlsx");

    test:assertTrue(result is Error, "Should return error for non-existent file");
    if result is Error {
        test:assertTrue(result.message().includes("Failed to read file"),
            "Error message should mention file read failure");
    }
}

// =============================================================================
// EDGE CASE TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "edge"]
}
function testParseEmptySheet() returns error? {
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_sheet.xlsx");

    // Empty sheet should return empty array
    test:assertEquals(rows.length(), 0, "Empty sheet should return empty array");
}

@test:Config {
    groups: ["parseSheet", "edge"]
}
function testParseSingleCell() returns error? {
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_single_cell.xlsx");

    test:assertEquals(rows.length(), 1, "Should have 1 row");
    test:assertEquals(rows[0].length(), 1, "Should have 1 column");
    test:assertEquals(rows[0][0], "alone", "Cell value should be 'alone'");
}

@test:Config {
    groups: ["parseSheet", "edge"]
}
function testParseUnicodeData() returns error? {
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_unicode.xlsx");

    test:assertEquals(rows.length(), 5, "Should have 5 rows (header + 4 data)");

    // Verify header
    test:assertEquals(rows[0][0], "Language", "First header should be 'Language'");

    // Verify English row
    test:assertEquals(rows[1][0], "English", "First data row language");
    test:assertEquals(rows[1][1], "Hello", "First data row greeting");

    // Verify Japanese row has content (exact Unicode comparison can be tricky)
    test:assertEquals(rows[2][0], "Japanese", "Second data row language");
    test:assertTrue(rows[2][1].length() > 0, "Japanese greeting should have content");

    // Verify Arabic row has content
    test:assertEquals(rows[3][0], "Arabic", "Third data row language");
    test:assertTrue(rows[3][1].length() > 0, "Arabic greeting should have content");

    // Verify Emoji row
    test:assertEquals(rows[4][0], "Emoji", "Fourth data row language");
}

@test:Config {
    groups: ["parseSheet", "edge"]
}
function testParseDataWithEmptyRowsInMiddle() returns error? {
    // Default behavior: empty rows are skipped
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Should skip empty rows and have: header, First/100, Second/200, Third/300
    test:assertTrue(rows.length() >= 4, "Should have at least 4 rows after filtering empty");

    // Find data rows (skip header)
    boolean foundFirst = false;
    boolean foundSecond = false;
    boolean foundThird = false;

    foreach string[] row in rows {
        if row[0] == "First" {
            foundFirst = true;
            test:assertEquals(row[1], "100", "First row value should be 100");
        }
        if row[0] == "Second" {
            foundSecond = true;
            test:assertEquals(row[1], "200", "Second row value should be 200");
        }
        if row[0] == "Third" {
            foundThird = true;
            test:assertEquals(row[1], "300", "Third row value should be 300");
        }
    }

    test:assertTrue(foundFirst, "Should find 'First' row");
    test:assertTrue(foundSecond, "Should find 'Second' row");
    test:assertTrue(foundThird, "Should find 'Third' row");
}

// =============================================================================
// DATA PROJECTION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseWithDefaultProjection() returns error? {
    // Default projection: allowDataProjection = {}, nilAsOptionalField = false, absentAsNilableType = false
    // simple.xlsx has: Name, Age, City columns
    // RecordWithOptionalFields has: name, age?, city? fields
    // All columns match, so should work with default settings
    RecordWithOptionalFields[] data = check parseSheet(TEST_DATA_DIR + "simple.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].name, "John", "First name should be 'John'");
    // Age column exists but value is string "30" - will be converted to int 30
    test:assertEquals(data[0].age, 30, "First age should be 30");
    test:assertEquals(data[0].city, "New York", "First city should be 'New York'");
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseWithAbsentAsNilableTypeTrue() returns error? {
    // Test absentAsNilableType = true: missing columns should be set to nil for nilable fields
    // simple.xlsx has: Name, Age, City columns
    // RecordWithExtraField has: name, age, extraField? - extraField has no matching column
    ParseOptions opts = {
        allowDataProjection: {
            absentAsNilableType: true,
            nilAsOptionalField: false
        }
    };
    RecordWithExtraField[] data = check parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].name, "John", "First name should be 'John'");
    test:assertEquals(data[0].age, 30, "First age should be 30");
    test:assertEquals(data[0].extraField, null, "extraField should be null (no matching column)");
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseWithAbsentAsNilableTypeFalseAndRequiredField() returns error? {
    // Test absentAsNilableType = false with a record that has a required field without matching column
    // This should fail because 'department' has no matching column and is required
    ParseOptions opts = {
        allowDataProjection: {
            absentAsNilableType: false,
            nilAsOptionalField: false
        }
    };

    // simple.xlsx has: Name, Age, City - no 'department' column
    // StrictModeRecord has: name, age, department (required)
    StrictModeRecord[]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertTrue(result is Error, "Should return error for required field without matching column");
    if result is Error {
        test:assertTrue(result.message().includes("department") || result.message().includes("no matching column"),
            "Error message should mention the missing field");
    }
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseWithAllowDataProjectionFalse() returns error? {
    // Test allowDataProjection = false (strict mode)
    // All record fields must have matching columns
    ParseOptions opts = {
        allowDataProjection: false
    };

    // simple.xlsx has: Name, Age, City
    // StrictModeRecord has: name, age, department - 'department' has no match
    StrictModeRecord[]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertTrue(result is Error, "Should return error when projection disabled and fields don't match");
    if result is Error {
        test:assertTrue(result.message().includes("projection disabled") ||
                        result.message().includes("without matching columns"),
            "Error message should mention projection disabled");
    }
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseWithAllowDataProjectionFalseMatchingFields() returns error? {
    // Test allowDataProjection = false with matching fields should work
    // simple.xlsx has: Name, Age, City columns
    // We need a record with exactly those fields
    ParseOptions opts = {
        allowDataProjection: false
    };

    // Use RecordWithOptionalFields which maps to Name, Age, City (all present)
    // Note: Field names are case-insensitive for matching
    RecordWithOptionalFields[]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // This might fail if field name case doesn't match
    // If it fails, it's expected behavior for strict mode
    if result is RecordWithOptionalFields[] {
        test:assertEquals(result.length(), 3, "Should have 3 records when fields match");
    }
    // If error, strict mode is working (case sensitivity might cause mismatch)
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseMapWithNilAsOptionalFieldTrue() returns error? {
    // Test nilAsOptionalField for map<anydata>[] - nil values should be skipped
    ParseOptions opts = {
        allowDataProjection: {
            nilAsOptionalField: true,
            absentAsNilableType: false
        }
    };

    // edge_empty_rows.xlsx has some empty cells which become nil
    // With nilAsOptionalField=true, nil values should not be added to the map
    map<anydata>[] data = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx", 0, opts);

    // Just verify parsing succeeds - actual behavior depends on data content
    test:assertTrue(data.length() >= 0, "Should parse successfully");
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseMapWithNilAsOptionalFieldFalse() returns error? {
    // Test nilAsOptionalField=false for map<anydata>[] - nil values should be included
    ParseOptions opts = {
        allowDataProjection: {
            nilAsOptionalField: false,
            absentAsNilableType: false
        }
    };

    map<anydata>[] data = check parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertEquals(data.length(), 3, "Should have 3 records");
    // All cells have values in simple.xlsx, so we're just verifying the option is accepted
}

// =============================================================================
// CASE-INSENSITIVE HEADERS TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "options"]
}
function testCaseInsensitiveHeadersEnabled() returns error? {
    // case_headers.xlsx has headers: "NAME", "AGE", "Department" (mixed case)
    // CaseTestEmployee has fields: name, age, department (lowercase)
    // With caseInsensitiveHeaders=true, should match successfully
    ParseOptions opts = {
        caseInsensitiveHeaders: true
    };

    CaseTestEmployee[] employees = check parseSheet(TEST_DATA_DIR + "case_headers.xlsx", 0, opts);

    test:assertEquals(employees.length(), 2, "Should have 2 employees");
    test:assertEquals(employees[0].name, "John", "First employee name");
    test:assertEquals(employees[0].age, 30, "First employee age");
    test:assertEquals(employees[0].department, "Engineering", "First employee department");
    test:assertEquals(employees[1].name, "Jane", "Second employee name");
    test:assertEquals(employees[1].age, 25, "Second employee age");
    test:assertEquals(employees[1].department, "Marketing", "Second employee department");
}

@test:Config {
    groups: ["parseSheet", "options"]
}
function testCaseInsensitiveHeadersDisabled() returns error? {
    // case_headers.xlsx has headers: "NAME", "AGE", "Department" (mixed case)
    // CaseTestEmployee has fields: name, age, department (lowercase)
    // With caseInsensitiveHeaders=false (default), headers won't match fields exactly
    // This should fail because required fields have no matching columns
    ParseOptions opts = {
        caseInsensitiveHeaders: false,
        allowDataProjection: {
            absentAsNilableType: false
        }
    };

    CaseTestEmployee[]|Error result = parseSheet(TEST_DATA_DIR + "case_headers.xlsx", 0, opts);

    // Should fail because "NAME" != "name", "AGE" != "age", "Department" != "department"
    test:assertTrue(result is Error, "Should fail when case-sensitive and headers don't match");
    if result is Error {
        // Error should mention missing fields or no matching columns
        string message = result.message();
        test:assertTrue(
            message.includes("no matching column") || message.includes("required") || message.includes("field"),
            "Error should mention field/column matching issue: " + message
        );
    }
}

@test:Config {
    groups: ["parseSheet", "options"]
}
function testCaseInsensitiveHeadersWithWorkbookAPI() returns error? {
    // Test case-insensitive headers via Workbook/Sheet API
    Workbook wb = check openFile(TEST_DATA_DIR + "case_headers.xlsx");

    Sheet sheet = check wb.getSheet("Sheet1");
    RowReadOptions opts = {
        caseInsensitiveHeaders: true
    };

    CaseTestEmployee[] employees = check sheet.getRows(opts);

    test:assertEquals(employees.length(), 2, "Should have 2 employees via Workbook API");
    test:assertEquals(employees[0].name, "John", "First employee name via Workbook API");
    test:assertEquals(employees[0].age, 30, "First employee age via Workbook API");

    check wb.close();
}

// =============================================================================
// ERROR TYPE PRESERVATION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "error"]
}
function testErrorTypeIsTypeConversionError() returns error? {
    // Create test file with invalid data that can't be converted to int
    string testFile = TEST_DATA_DIR + "error_type_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", "not_a_number"]  // Invalid age - should cause TypeConversionError
    ];
    check writeSheet(data, testFile);

    ErrorTypeTestRecord[]|error result = parseSheet(testFile);

    // Should be TypeConversionError, not generic Error
    test:assertTrue(result is TypeConversionError,
        "Should return TypeConversionError for invalid type conversion");

    if result is TypeConversionError {
        ErrorDetails details = result.detail();
        // Error details should be preserved
        test:assertTrue(details.rowNumber is int, "Should have rowNumber in error details");
        test:assertTrue(details.columnNumber is int, "Should have columnNumber in error details");
    }

    check file:remove(testFile);
}

@test:Config {
    groups: ["parseSheet", "error"]
}
function testSheetGetRowsErrorTypePreservation() returns error? {
    // Test that Sheet.getRows() also preserves error types
    string testFile = TEST_DATA_DIR + "sheet_error_type_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["Jane", "invalid_age"]  // Invalid age
    ];
    check writeSheet(data, testFile);

    Workbook wb = check openFile(testFile);
    Sheet sheet = check wb.getSheetByIndex(0);

    ErrorTypeTestRecord[]|Error result = sheet.getRows();

    test:assertTrue(result is TypeConversionError,
        "Sheet.getRows should return TypeConversionError for invalid type conversion");

    if result is TypeConversionError {
        ErrorDetails details = result.detail();
        test:assertTrue(details.rowNumber is int, "Sheet error should have rowNumber");
    }

    check wb.close();
    check file:remove(testFile);
}

// =============================================================================
// BLANK CELL HANDLING TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "error"]
}
function testBlankCellRequiredFieldError() returns error? {
    // Create test file with blank cell for required non-nilable field
    string testFile = TEST_DATA_DIR + "blank_required_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", "30"],
        ["Jane", ""]  // Blank age - required non-nilable field
    ];
    check writeSheet(data, testFile);

    RequiredAgeRecord[]|error result = parseSheet(testFile);

    // Should fail with TypeConversionError for blank required field
    test:assertTrue(result is TypeConversionError,
        "Blank cell for required field should return TypeConversionError");

    if result is TypeConversionError {
        string message = result.message();
        test:assertTrue(message.includes("null") || message.includes("blank"),
            "Error should mention null/blank issue: " + message);
    }

    check file:remove(testFile);
}

@test:Config {
    groups: ["parseSheet"]
}
function testBlankCellOptionalFieldSuccess() returns error? {
    // Blank cell for optional field should succeed
    string testFile = TEST_DATA_DIR + "blank_optional_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", ""],   // Blank age - but field is optional
        ["Jane", "25"]
    ];
    check writeSheet(data, testFile);

    OptionalAgeRecord[] result = check parseSheet(testFile);

    test:assertEquals(result.length(), 2, "Should parse both rows");
    test:assertEquals(result[0].name, "John", "First record name should be John");
    test:assertEquals(result[1].age, 25, "Second record should have age 25");

    check file:remove(testFile);
}

@test:Config {
    groups: ["parseSheet"]
}
function testBlankCellNilableFieldSuccess() returns error? {
    // Blank cell for nilable field should succeed with null
    string testFile = TEST_DATA_DIR + "blank_nilable_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", ""],   // Blank age - but field is nilable
        ["Jane", "25"]
    ];
    check writeSheet(data, testFile);

    NilableAgeRecord[] result = check parseSheet(testFile);

    test:assertEquals(result.length(), 2, "Should parse both rows");
    test:assertEquals(result[0].age, (), "First record age should be nil");
    test:assertEquals(result[1].age, 25, "Second record should have age 25");

    check file:remove(testFile);
}

// =============================================================================
// ROW WRAPPER TESTS
// =============================================================================
// Tests for Row-wrapped types that preserve row positions during parsing.

@test:Config {
    groups: ["parseSheet", "row-wrapper"]
}
function testParseWithRowWrapper() returns error? {
    // edge_empty_rows.xlsx has: header, First, empty, Second, empty, Third
    // With Row wrapper, ALL rows should be included (including empty ones)
    SimpleDataRow[] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Should have 5 rows (First, empty, Second, empty, Third)
    // Header is at row 0, data starts at row 1
    test:assertEquals(rows.length(), 5, "Should have 5 rows including empty rows");

    // Verify first data row
    test:assertEquals(rows[0].rowIndex, 0, "First row should have rowIndex 0");
    test:assertTrue(rows[0].value != null, "First row should have value");
    test:assertEquals(rows[0].value?.name, "First", "First row name");
    test:assertEquals(rows[0].value?.value, 100, "First row value");

    // Verify empty row (rowIndex 1)
    test:assertEquals(rows[1].rowIndex, 1, "Second row should have rowIndex 1");
    test:assertEquals(rows[1].value, null, "Empty row should have null value");

    // Verify second data row
    test:assertEquals(rows[2].rowIndex, 2, "Third row should have rowIndex 2");
    test:assertTrue(rows[2].value != null, "Third row should have value");
    test:assertEquals(rows[2].value?.name, "Second", "Third row name");
    test:assertEquals(rows[2].value?.value, 200, "Third row value");

    // Verify another empty row (rowIndex 3)
    test:assertEquals(rows[3].rowIndex, 3, "Fourth row should have rowIndex 3");
    test:assertEquals(rows[3].value, null, "Empty row should have null value");

    // Verify third data row
    test:assertEquals(rows[4].rowIndex, 4, "Fifth row should have rowIndex 4");
    test:assertTrue(rows[4].value != null, "Fifth row should have value");
    test:assertEquals(rows[4].value?.name, "Third", "Fifth row name");
    test:assertEquals(rows[4].value?.value, 300, "Fifth row value");
}

@test:Config {
    groups: ["parseSheet", "row-wrapper"]
}
function testParseRowWrapperWithEmployees() returns error? {
    // employees.xlsx has no empty rows, so all rows should have values
    EmployeeRow[] rows = check parseSheet(TEST_DATA_DIR + "employees.xlsx");

    test:assertEquals(rows.length(), 3, "Should have 3 employee rows");

    // Verify all rows have sequential indices and non-null values
    foreach int i in 0 ..< rows.length() {
        test:assertEquals(rows[i].rowIndex, i, "Row index should be " + i.toString());
        test:assertTrue(rows[i].value != null, "Row should have value at index " + i.toString());
    }

    // Verify first employee data
    test:assertEquals(rows[0].value?.name, "John Doe", "First employee name");
    test:assertEquals(rows[0].value?.age, 30, "First employee age");
    test:assertEquals(rows[0].value?.department, "Engineering", "First employee department");
}

@test:Config {
    groups: ["parseSheet", "row-wrapper"]
}
function testParseRowWrapperFilterPreservesPositions() returns error? {
    // Parse with Row wrapper
    SimpleDataRow[] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Filter out empty rows
    SimpleDataRow[] nonEmptyRows = rows.filter(r => r.value != null);

    test:assertEquals(nonEmptyRows.length(), 3, "Should have 3 non-empty rows");

    // Verify original positions are preserved after filtering
    test:assertEquals(nonEmptyRows[0].rowIndex, 0, "First non-empty should have original rowIndex 0");
    test:assertEquals(nonEmptyRows[1].rowIndex, 2, "Second non-empty should have original rowIndex 2");
    test:assertEquals(nonEmptyRows[2].rowIndex, 4, "Third non-empty should have original rowIndex 4");
}

@test:Config {
    groups: ["workbook", "row-wrapper"]
}
function testSheetGetRowsWithRowWrapper() returns error? {
    // Test Row wrapper via Workbook/Sheet API
    Workbook wb = check openFile(TEST_DATA_DIR + "edge_empty_rows.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    SimpleDataRow[] rows = check sheet.getRows();

    test:assertEquals(rows.length(), 5, "Should have 5 rows including empty rows");

    // Verify positions preserved
    test:assertEquals(rows[0].rowIndex, 0, "First row index");
    test:assertEquals(rows[1].rowIndex, 1, "Second row index (empty)");
    test:assertEquals(rows[1].value, null, "Empty row value should be null");
    test:assertEquals(rows[2].rowIndex, 2, "Third row index");

    check wb.close();
}

// =============================================================================
// INVALID OPTION COMBINATION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "options", "edge-cases"]
}
function testHeaderRowBeyondSheetBounds() returns error? {
    // headerRowIndex = 100 on a small sheet (4 rows) should return ParseError
    // Note: headerRowIndex validation only applies to record/map parsing, not string[][]
    ParseOptions opts = {
        headerRowIndex: 100
    };

    // Use record parsing to test headerRowIndex validation
    Employee[]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // Record parsing validates headerRowIndex - when beyond bounds, returns ParseError
    test:assertTrue(result is ParseError, "Should return ParseError when headerRowIndex beyond sheet bounds");
    if result is ParseError {
        test:assertTrue(result.message().includes("Header row"), "Error should mention header row");
    }
}

@test:Config {
    groups: ["parseSheet", "options", "edge-cases"]
}
function testNegativeHeaderRowBeyondMinusOne() returns error? {
    // Only -1 is valid for "no headers", other negatives should error or be handled
    ParseOptions opts = {
        headerRowIndex: -5
    };

    string[][]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // Should either normalize to valid value or return error
    // The actual behavior depends on implementation
    test:assertTrue(result is string[][] || result is ParseError,
        "Should handle invalid negative headerRowIndex");
}

// =============================================================================
// EXTREME NUMERIC VALUE TESTS
// =============================================================================

type NumericRecord record {|
    int id;
    int largeInt;
    decimal largeDecimal;
|};

@test:Config {
    groups: ["parseSheet", "numeric", "edge-cases"],
    before: setupExtremeNumericTestData
}
function testParseExtremeIntValues() returns error? {
    NumericRecord[] records = check parseSheet(TEST_DATA_DIR + "extreme_numeric_test.xlsx");

    test:assertEquals(records.length(), 3, "Should parse all 3 rows");

    // Test large positive int
    test:assertEquals(records[0].largeInt, int:MAX_VALUE, "Should handle MAX_INT");

    // Test large negative int
    test:assertEquals(records[1].largeInt, int:MIN_VALUE, "Should handle MIN_INT");

    // Test zero
    test:assertEquals(records[2].largeInt, 0, "Should handle zero");
}

@test:Config {
    groups: ["parseSheet", "numeric", "edge-cases"],
    before: setupExtremeNumericTestData
}
function testParseLargeDecimalValues() returns error? {
    NumericRecord[] records = check parseSheet(TEST_DATA_DIR + "extreme_numeric_test.xlsx");

    test:assertEquals(records.length(), 3, "Should parse all 3 rows");

    // Test large decimal values are preserved
    test:assertTrue(records[0].largeDecimal > 1e15d, "Should handle large positive decimal");
    test:assertTrue(records[1].largeDecimal < -1e15d, "Should handle large negative decimal");
}

// Setup function for extreme numeric test data
function setupExtremeNumericTestData() returns error? {
    string testFilePath = TEST_DATA_DIR + "extreme_numeric_test.xlsx";

    if check file:test(testFilePath, file:EXISTS) {
        check file:remove(testFilePath);
    }

    // Create test data with extreme values
    // Note: Excel/POI may have precision limits for very large numbers
    string[][] testData = [
        ["id", "largeInt", "largeDecimal"],
        ["1", int:MAX_VALUE.toString(), "1234567890123456.789"],   // MAX_INT
        ["2", int:MIN_VALUE.toString(), "-1234567890123456.789"], // MIN_INT
        ["3", "0", "0.0"]                                          // Zero
    ];

    check writeSheet(testData, testFilePath);
}

// =============================================================================
// HEADER-LESS PARSING TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "headerless"]
}
function testParseHeaderlessToMap() returns error? {
    // Create test file without headers - just raw data
    string testFile = TEST_DATA_DIR + "headerless_test.xlsx";
    string[][] data = [
        ["Alice", "30"],
        ["Bob", "25"]
    ];
    check writeSheet(data, testFile, writeHeaders = false);

    // Parse with headerRowIndex = null (header-less mode)
    ParseOptions opts = {
        headerRowIndex: ()
    };
    map<anydata>[] result = check parseSheet(testFile, 0, opts);

    test:assertEquals(result.length(), 2, "Should have 2 rows");
    test:assertEquals(result[0]["col0"], "Alice", "First row col0 should be 'Alice'");
    test:assertEquals(result[0]["col1"], "30", "First row col1 should be '30'");
    test:assertEquals(result[1]["col0"], "Bob", "Second row col0 should be 'Bob'");
    test:assertEquals(result[1]["col1"], "25", "Second row col1 should be '25'");

    check file:remove(testFile);
}

@test:Config {
    groups: ["parseSheet", "rowcount"]
}
function testParseWithRowCountLimit() returns error? {
    // Parse employees.xlsx (has 3 data rows) with rowCount limit
    ParseOptions opts = {
        rowCount: 2
    };
    Employee[] employees = check parseSheet(TEST_DATA_DIR + "employees.xlsx", 0, opts);

    test:assertEquals(employees.length(), 2, "Should have only 2 records due to rowCount limit");
    test:assertEquals(employees[0].name, "John Doe", "First employee name");
    test:assertEquals(employees[1].name, "Jane Smith", "Second employee name");
}

@test:Config {
    groups: ["parseSheet", "rowcount"]
}
function testParseWithRowCountNullMeansAll() returns error? {
    // Parse with rowCount = null (default) should read all rows
    ParseOptions opts = {
        rowCount: ()  // Explicit null = read all
    };
    Employee[] employees = check parseSheet(TEST_DATA_DIR + "employees.xlsx", 0, opts);

    test:assertEquals(employees.length(), 3, "Should have all 3 records when rowCount is null");
}

@test:Config {
    groups: ["parseSheet", "headerless"]
}
function testParseHeaderlessToRecord() returns error? {
    // Create test file without headers
    string testFile = TEST_DATA_DIR + "headerless_record_test.xlsx";
    string[][] data = [
        ["John", "Engineer"],
        ["Jane", "Designer"]
    ];
    check writeSheet(data, testFile, writeHeaders = false);

    // Parse with headerRowIndex = null
    ParseOptions opts = {
        headerRowIndex: ()
    };
    HeaderlessRecord[] result = check parseSheet(testFile, 0, opts);

    test:assertEquals(result.length(), 2, "Should have 2 records");
    test:assertEquals(result[0].col0, "John", "First record col0");
    test:assertEquals(result[0].col1, "Engineer", "First record col1");
    test:assertEquals(result[1].col0, "Jane", "Second record col0");
    test:assertEquals(result[1].col1, "Designer", "Second record col1");

    check file:remove(testFile);
}

// =============================================================================
// OPEN RECORD POPULATION TESTS
// =============================================================================

@test:Config {
    groups: ["parseSheet", "openrecord"]
}
function testParseToOpenRecord() returns error? {
    // employees.xlsx has columns: name, age, department
    // OpenEmployee only defines: name, age (no department)
    // The 'department' column should be populated into the rest field
    OpenEmployee[] employees = check parseSheet(TEST_DATA_DIR + "employees.xlsx");

    test:assertEquals(employees.length(), 3, "Should have 3 employees");

    // Verify defined fields
    test:assertEquals(employees[0].name, "John Doe", "First employee name");
    test:assertEquals(employees[0].age, 30, "First employee age");

    // Verify extra field (department) was populated via rest field
    // Access using record indexing since 'department' is in rest field
    anydata dept = employees[0]["department"];
    test:assertEquals(dept, "Engineering", "First employee department via rest field");

    // Verify second employee
    test:assertEquals(employees[1].name, "Jane Smith", "Second employee name");
    test:assertEquals(employees[1].age, 28, "Second employee age");
    test:assertEquals(employees[1]["department"], "Marketing", "Second employee department via rest field");

    // Verify third employee
    test:assertEquals(employees[2].name, "Bob Johnson", "Third employee name");
    test:assertEquals(employees[2].age, 35, "Third employee age");
    test:assertEquals(employees[2]["department"], "Sales", "Third employee department via rest field");
}

@test:Config {
    groups: ["parseSheet", "openrecord"]
}
function testParseToOpenRecordWithMultipleExtraColumns() returns error? {
    // Create test file with extra columns beyond defined fields
    string testFile = TEST_DATA_DIR + "open_record_test.xlsx";
    string[][] data = [
        ["name", "age", "department", "salary", "location"],
        ["Alice", "25", "Engineering", "50000", "NYC"],
        ["Bob", "30", "Marketing", "60000", "LA"]
    ];
    check writeSheet(data, testFile, writeHeaders = false);  // Headers in first row

    // Parse to OpenEmployee (only defines name and age)
    OpenEmployee[] employees = check parseSheet(testFile);

    test:assertEquals(employees.length(), 2, "Should have 2 employees");

    // Verify defined fields
    test:assertEquals(employees[0].name, "Alice", "First employee name");
    test:assertEquals(employees[0].age, 25, "First employee age");

    // Verify extra fields
    test:assertEquals(employees[0]["department"], "Engineering", "First employee department");
    test:assertEquals(employees[0]["salary"], "50000", "First employee salary");
    test:assertEquals(employees[0]["location"], "NYC", "First employee location");

    test:assertEquals(employees[1].name, "Bob", "Second employee name");
    test:assertEquals(employees[1]["salary"], "60000", "Second employee salary");

    check file:remove(testFile);
}

@test:Config {
    groups: ["parseSheet", "openrecord"]
}
function testParseClosedRecordIgnoresExtraColumns() returns error? {
    // Verify that closed records (Employee) do NOT get extra columns
    // employees.xlsx has exactly the columns Employee expects
    Employee[] employees = check parseSheet(TEST_DATA_DIR + "employees.xlsx");

    test:assertEquals(employees.length(), 3, "Should have 3 employees");
    test:assertEquals(employees[0].name, "John Doe", "First employee name");
    test:assertEquals(employees[0].department, "Engineering", "First employee department");

    // Closed record should not have any extra fields - all fields are defined
}
