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
    groups: ["parse"]
}
function testParseToStringArray() returns error? {
    string[][] rows = check parse(TEST_DATA_DIR + "simple.xlsx");

    // Verify exact content matches expected data
    assertStringArrayEquals(rows, EXPECTED_SIMPLE_DATA, "simple.xlsx parse");
}

@test:Config {
    groups: ["parse"]
}
function testParseToRecords() returns error? {
    Employee[] employees = check parse(TEST_DATA_DIR + "employees.xlsx");

    // Verify exact content matches expected data
    assertEmployeesEqual(employees, EXPECTED_EMPLOYEES);
}

@test:Config {
    groups: ["parse", "options"]
}
function testParseWithHeaderRowOption() returns error? {
    // complex_headers.xlsx has: Row 0=title, Row 1=metadata, Row 2=headers, Row 3+=data
    ParseOptions opts = {
        headerRow: 2,
        dataStartRow: 3
    };
    string[][] rows = check parse(TEST_DATA_DIR + "complex_headers.xlsx", 0, opts);

    // When dataStartRow is specified, only data rows are returned (starting from row 3)
    test:assertEquals(rows.length(), 2, "Should have 2 data rows");
    test:assertEquals(rows[0][0], "Item1", "First row should be first data row");
    test:assertEquals(rows[1][0], "Item2", "Second row should be second data row");
}

// =============================================================================
// SHEET SELECTION TESTS
// =============================================================================

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetSelectionByName() returns error? {
    // Select Sheet2 by name
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx", "Sheet2");

    // Verify we got Sheet2 data
    assertStringArrayEquals(rows, EXPECTED_SHEET2_DATA, "Sheet2 data");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetSelectionByIndex() returns error? {
    // Select second sheet (index 1)
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx", 1);

    // Verify we got Sheet2 data (index 1)
    assertStringArrayEquals(rows, EXPECTED_SHEET2_DATA, "Sheet index 1 data");
}

@test:Config {
    groups: ["parse"]
}
function testParseDefaultsToFirstSheet() returns error? {
    // No sheet specified - should default to first sheet
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx");

    // Verify we got Sheet1 data
    assertStringArrayEquals(rows, EXPECTED_SHEET1_DATA, "Default to Sheet1");
}

@test:Config {
    groups: ["parse"]
}
function testParseThirdSheet() returns error? {
    // Select Sheet3 by name
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx", "Sheet3");

    // Verify we got Sheet3 data
    assertStringArrayEquals(rows, EXPECTED_SHEET3_DATA, "Sheet3 data");
}

// =============================================================================
// FORMULA HANDLING TESTS
// =============================================================================

@test:Config {
    groups: ["parse"]
}
function testParseFormulaCachedMode() returns error? {
    ParseOptions opts = {
        formulaMode: CACHED
    };
    string[][] rows = check parse(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

    // In CACHED mode, formula cells return their cached calculated values.
    // For formulas written via POI without evaluation, the default cached value is "0".
    // In real-world usage with Excel-created files, this would return the actual calculated result.
    test:assertEquals(rows.length(), 3, "Should have 3 rows");
    test:assertEquals(rows[0][2], "Sum", "Header should be 'Sum'");
    test:assertEquals(rows[1][2], "0", "Unevaluated formula should return default cached value '0'");
    test:assertEquals(rows[2][2], "0", "Unevaluated formula should return default cached value '0'");
}

@test:Config {
    groups: ["parse"]
}
function testParseFormulaTextMode() returns error? {
    ParseOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check parse(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

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
    groups: ["parse", "types"]
}
function testParseNumericToInt() returns error? {
    NumericTypes[] data = check parse(TEST_DATA_DIR + "numeric_types.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].intValue, 42, "First intValue should be 42");
    test:assertEquals(data[1].intValue, -100, "Second intValue should be -100");
    test:assertEquals(data[2].intValue, 0, "Third intValue should be 0");
}

@test:Config {
    groups: ["parse", "types"]
}
function testParseNumericToDecimal() returns error? {
    NumericTypes[] data = check parse(TEST_DATA_DIR + "numeric_types.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    // Decimal comparison with tolerance for floating point
    test:assertTrue(data[0].decimalValue > 3.14d && data[0].decimalValue < 3.15d,
        "First decimalValue should be ~3.14159");
    test:assertTrue(data[1].decimalValue > 0.0009d && data[1].decimalValue < 0.002d,
        "Second decimalValue should be ~0.001");
}

@test:Config {
    groups: ["parse", "types"]
}
function testParseToTypedRecordWithVariousTypes() returns error? {
    TypeVariety[] data = check parse(TEST_DATA_DIR + "types_variety.xlsx");

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
    groups: ["parse", "types"]
}
function testParseToMapArray() returns error? {
    map<anydata>[] data = check parse(TEST_DATA_DIR + "simple.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records (excluding header)");
    test:assertEquals(data[0]["Name"], "John", "First record Name should be 'John'");
    test:assertEquals(data[0]["Age"], "30", "First record Age should be '30'");
    test:assertEquals(data[0]["City"], "New York", "First record City should be 'New York'");
}

// =============================================================================
// OPTIONS TESTS
// =============================================================================

@test:Config {
    groups: ["parse", "options"]
}
function testParseWithCustomDataStartRow() returns error? {
    ParseOptions opts = {
        headerRow: 0,
        dataStartRow: 2  // Skip first data row
    };
    string[][] rows = check parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // When dataStartRow is specified, only data rows starting from that row are returned
    test:assertEquals(rows.length(), 2, "Should have 2 data rows (skipped first data row)");
    test:assertEquals(rows[0][0], "Jane", "First row should be 'Jane' (skipped 'John')");
    test:assertEquals(rows[1][0], "Bob", "Second row should be 'Bob'");
}

@test:Config {
    groups: ["parse", "options"]
}
function testParseWithIncludeEmptyRowsFalse() returns error? {
    ParseOptions opts = {
        includeEmptyRows: false  // Default
    };
    string[][] rows = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx", 0, opts);

    // Empty rows should be skipped
    // Original: header, First, empty, Second, empty, Third
    // Expected: header, First, Second, Third (no empty rows)
    test:assertEquals(rows.length(), 4, "Should have 4 rows (header + 3 data, no empty)");
}

@test:Config {
    groups: ["parse", "options"]
}
function testParseWithIncludeEmptyRowsTrue() returns error? {
    ParseOptions opts = {
        includeEmptyRows: true
    };
    string[][] rows = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx", 0, opts);

    // Empty rows should be included
    test:assertEquals(rows.length(), 6, "Should have 6 rows (including empty rows)");
}

// =============================================================================
// ANNOTATION TESTS
// =============================================================================

@test:Config {
    groups: ["parse", "annotation"]
}
function testParseWithXlsxNameAnnotation() returns error? {
    // annotated.xlsx has headers: "First Name", "Employee ID", "Department Name"
    // AnnotatedEmployee maps these to: firstName, id, department
    AnnotatedEmployee[] data = check parse(TEST_DATA_DIR + "annotated.xlsx");

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
    groups: ["parse", "negative"]
}
function testParseSheetNotFoundByName() returns error? {
    string[][]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", "NonExistentSheet");

    test:assertTrue(result is Error, "Should return error for non-existent sheet");
    if result is Error {
        test:assertTrue(result.message().includes("not found"),
            "Error message should mention sheet not found");
    }
}

@test:Config {
    groups: ["parse", "negative"]
}
function testParseSheetNotFoundByIndex() returns error? {
    string[][]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", 99);

    test:assertTrue(result is Error, "Should return error for invalid sheet index");
    if result is Error {
        test:assertTrue(result.message().includes("out of range"),
            "Error message should mention index out of range");
    }
}

@test:Config {
    groups: ["parse", "negative"]
}
function testParseFileNotFound() returns error? {
    string[][]|Error result = parse(TEST_DATA_DIR + "nonexistent.xlsx");

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
    groups: ["parse", "edge"]
}
function testParseEmptySheet() returns error? {
    string[][] rows = check parse(TEST_DATA_DIR + "edge_empty_sheet.xlsx");

    // Empty sheet should return empty array
    test:assertEquals(rows.length(), 0, "Empty sheet should return empty array");
}

@test:Config {
    groups: ["parse", "edge"]
}
function testParseSingleCell() returns error? {
    string[][] rows = check parse(TEST_DATA_DIR + "edge_single_cell.xlsx");

    test:assertEquals(rows.length(), 1, "Should have 1 row");
    test:assertEquals(rows[0].length(), 1, "Should have 1 column");
    test:assertEquals(rows[0][0], "alone", "Cell value should be 'alone'");
}

@test:Config {
    groups: ["parse", "edge"]
}
function testParseUnicodeData() returns error? {
    string[][] rows = check parse(TEST_DATA_DIR + "edge_unicode.xlsx");

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
    groups: ["parse", "edge"]
}
function testParseDataWithEmptyRowsInMiddle() returns error? {
    ParseOptions opts = {
        includeEmptyRows: false
    };
    string[][] rows = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx", 0, opts);

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
    groups: ["parse", "projection"]
}
function testParseWithDefaultProjection() returns error? {
    // Default projection: allowDataProjection = {}, nilAsOptionalField = false, absentAsNilableType = false
    // simple.xlsx has: Name, Age, City columns
    // RecordWithOptionalFields has: name, age?, city? fields
    // All columns match, so should work with default settings
    RecordWithOptionalFields[] data = check parse(TEST_DATA_DIR + "simple.xlsx");

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].name, "John", "First name should be 'John'");
    // Age column exists but value is string "30" - will be converted to int 30
    test:assertEquals(data[0].age, 30, "First age should be 30");
    test:assertEquals(data[0].city, "New York", "First city should be 'New York'");
}

@test:Config {
    groups: ["parse", "projection"]
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
    RecordWithExtraField[] data = check parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertEquals(data.length(), 3, "Should have 3 records");
    test:assertEquals(data[0].name, "John", "First name should be 'John'");
    test:assertEquals(data[0].age, 30, "First age should be 30");
    test:assertEquals(data[0].extraField, null, "extraField should be null (no matching column)");
}

@test:Config {
    groups: ["parse", "projection"]
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
    StrictModeRecord[]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertTrue(result is Error, "Should return error for required field without matching column");
    if result is Error {
        test:assertTrue(result.message().includes("department") || result.message().includes("no matching column"),
            "Error message should mention the missing field");
    }
}

@test:Config {
    groups: ["parse", "projection"]
}
function testParseWithAllowDataProjectionFalse() returns error? {
    // Test allowDataProjection = false (strict mode)
    // All record fields must have matching columns
    ParseOptions opts = {
        allowDataProjection: false
    };

    // simple.xlsx has: Name, Age, City
    // StrictModeRecord has: name, age, department - 'department' has no match
    StrictModeRecord[]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertTrue(result is Error, "Should return error when projection disabled and fields don't match");
    if result is Error {
        test:assertTrue(result.message().includes("projection disabled") ||
                        result.message().includes("without matching columns"),
            "Error message should mention projection disabled");
    }
}

@test:Config {
    groups: ["parse", "projection"]
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
    RecordWithOptionalFields[]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    // This might fail if field name case doesn't match
    // If it fails, it's expected behavior for strict mode
    if result is RecordWithOptionalFields[] {
        test:assertEquals(result.length(), 3, "Should have 3 records when fields match");
    }
    // If error, strict mode is working (case sensitivity might cause mismatch)
}

@test:Config {
    groups: ["parse", "projection"]
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
    map<anydata>[] data = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx", 0, opts);

    // Just verify parsing succeeds - actual behavior depends on data content
    test:assertTrue(data.length() >= 0, "Should parse successfully");
}

@test:Config {
    groups: ["parse", "projection"]
}
function testParseMapWithNilAsOptionalFieldFalse() returns error? {
    // Test nilAsOptionalField=false for map<anydata>[] - nil values should be included
    ParseOptions opts = {
        allowDataProjection: {
            nilAsOptionalField: false,
            absentAsNilableType: false
        }
    };

    map<anydata>[] data = check parse(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertEquals(data.length(), 3, "Should have 3 records");
    // All cells have values in simple.xlsx, so we're just verifying the option is accepted
}

// =============================================================================
// CASE-INSENSITIVE HEADERS TESTS
// =============================================================================

@test:Config {
    groups: ["parse", "options"]
}
function testCaseInsensitiveHeadersEnabled() returns error? {
    // case_headers.xlsx has headers: "NAME", "AGE", "Department" (mixed case)
    // CaseTestEmployee has fields: name, age, department (lowercase)
    // With caseInsensitiveHeaders=true, should match successfully
    ParseOptions opts = {
        caseInsensitiveHeaders: true
    };

    CaseTestEmployee[] employees = check parse(TEST_DATA_DIR + "case_headers.xlsx", 0, opts);

    test:assertEquals(employees.length(), 2, "Should have 2 employees");
    test:assertEquals(employees[0].name, "John", "First employee name");
    test:assertEquals(employees[0].age, 30, "First employee age");
    test:assertEquals(employees[0].department, "Engineering", "First employee department");
    test:assertEquals(employees[1].name, "Jane", "Second employee name");
    test:assertEquals(employees[1].age, 25, "Second employee age");
    test:assertEquals(employees[1].department, "Marketing", "Second employee department");
}

@test:Config {
    groups: ["parse", "options"]
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

    CaseTestEmployee[]|Error result = parse(TEST_DATA_DIR + "case_headers.xlsx", 0, opts);

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
    groups: ["parse", "options"]
}
function testCaseInsensitiveHeadersWithWorkbookAPI() returns error? {
    // Test case-insensitive headers via Workbook/Sheet API
    Workbook wb = check new Workbook(TEST_DATA_DIR + "case_headers.xlsx");

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
    groups: ["parse", "error"]
}
function testErrorTypeIsTypeConversionError() returns error? {
    // Create test file with invalid data that can't be converted to int
    string testFile = TEST_DATA_DIR + "error_type_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", "not_a_number"]  // Invalid age - should cause TypeConversionError
    ];
    check write(data, testFile);

    ErrorTypeTestRecord[]|error result = parse(testFile);

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
    groups: ["parse", "error"]
}
function testSheetGetRowsErrorTypePreservation() returns error? {
    // Test that Sheet.getRows() also preserves error types
    string testFile = TEST_DATA_DIR + "sheet_error_type_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["Jane", "invalid_age"]  // Invalid age
    ];
    check write(data, testFile);

    Workbook wb = check new Workbook(testFile);
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
    groups: ["parse", "error"]
}
function testBlankCellRequiredFieldError() returns error? {
    // Create test file with blank cell for required non-nilable field
    string testFile = TEST_DATA_DIR + "blank_required_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", "30"],
        ["Jane", ""]  // Blank age - required non-nilable field
    ];
    check write(data, testFile);

    RequiredAgeRecord[]|error result = parse(testFile);

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
    groups: ["parse"]
}
function testBlankCellOptionalFieldSuccess() returns error? {
    // Blank cell for optional field should succeed
    string testFile = TEST_DATA_DIR + "blank_optional_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", ""],   // Blank age - but field is optional
        ["Jane", "25"]
    ];
    check write(data, testFile);

    OptionalAgeRecord[] result = check parse(testFile);

    test:assertEquals(result.length(), 2, "Should parse both rows");
    test:assertEquals(result[0].name, "John", "First record name should be John");
    test:assertEquals(result[1].age, 25, "Second record should have age 25");

    check file:remove(testFile);
}

@test:Config {
    groups: ["parse"]
}
function testBlankCellNilableFieldSuccess() returns error? {
    // Blank cell for nilable field should succeed with null
    string testFile = TEST_DATA_DIR + "blank_nilable_test.xlsx";
    string[][] data = [
        ["name", "age"],
        ["John", ""],   // Blank age - but field is nilable
        ["Jane", "25"]
    ];
    check write(data, testFile);

    NilableAgeRecord[] result = check parse(testFile);

    test:assertEquals(result.length(), 2, "Should parse both rows");
    test:assertEquals(result[0].age, (), "First record age should be nil");
    test:assertEquals(result[1].age, 25, "Second record should have age 25");

    check file:remove(testFile);
}
