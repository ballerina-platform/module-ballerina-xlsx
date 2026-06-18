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

    // CACHED mode returns each formula cell's stored result. The fixture carries the
    // values an Excel-authored file would hold (=A2+B2 → 30, =A3+B3 → 40).
    assertStringArrayEquals(rows, EXPECTED_FORMULA_CACHED, "formula cached mode");
}

@test:Config {
    groups: ["parseSheet"]
}
function testParseFormulaTextMode() returns error? {
    ParseOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check parseSheet(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

    // In TEXT mode, formula cells return the formula string with the "=" prefix.
    assertStringArrayEquals(rows, EXPECTED_FORMULA_TEXT, "formula text mode");
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
    map<CellValue>[] data = check parseSheet(TEST_DATA_DIR + "simple.xlsx");

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
function testParsePreservesEmptyRows() returns error? {
    // Every row in the used range produces an entry — empty
    // rows are NOT filtered. Empty cells are bound by the standard cell-binding
    // path (string[][] → padded empty strings; record → nil for nilable fields).
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // File contents: header, First, empty, Second, empty, Third
    // All 6 rows are returned (no skipping).
    test:assertEquals(rows.length(), 6, "Should have 6 rows (header + 3 data + 2 empty)");
}

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
    // all rows in the used range come through (including empty).
    // This test locates data rows by content rather than by index, so it works
    // regardless of whether empty rows are present in the output.
    string[][] rows = check parseSheet(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // File contents: header, First/100, empty, Second/200, empty, Third/300
    test:assertTrue(rows.length() >= 4, "Should have at least 4 rows (data is found by name below)");

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

    // Row order and the interleaved empty rows must be preserved exactly:
    // header(0), First(1), empty(2), Second(3), empty(4), Third(5).
    test:assertEquals(rows.length(), 6, "All 6 rows (header + 3 data + 2 empty) must come through");
    test:assertEquals(rows[1][0], "First", "Data row 'First' must be at index 1");
    test:assertEquals(rows[2][0], "", "Empty row must be preserved at index 2");
    test:assertEquals(rows[3][0], "Second", "Data row 'Second' must be at index 3");
    test:assertEquals(rows[4][0], "", "Empty row must be preserved at index 4");
    test:assertEquals(rows[5][0], "Third", "Data row 'Third' must be at index 5");
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

    // RecordWithOptionalFields uses @Name annotations mapping to "Name", "Age",
    // "City" — exactly the headers in simple.xlsx. With every field matched, strict
    // mode (allowDataProjection: false) must succeed and return all data rows.
    RecordWithOptionalFields[] result = check parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

    test:assertEquals(result.length(), 3, "Should have 3 records when every field matches a column");
    test:assertEquals(result[0].name, "John", "First record name");
    test:assertEquals(result[0].age, 30, "First record age");
    test:assertEquals(result[0].city, "New York", "First record city");
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseMapWithNilAsOptionalFieldTrue() returns error? {
    // Test nilAsOptionalField for map<CellValue>[] - nil values should be skipped
    ParseOptions opts = {
        allowDataProjection: {
            nilAsOptionalField: true,
            absentAsNilableType: false
        }
    };

    // Build a fixture with a GENUINELY blank cell (a skipped cell, not an empty
    // string) so the read hits the BLANK → nil path. With nilAsOptionalField=true,
    // that column's key must NOT be present in the map at all.
    string nilMapFile = TEST_DATA_DIR + "temp_nil_skip_map.xlsx";
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, "name");
    check sheet.setCell(0, 1, "department");
    check sheet.setCell(1, 0, "Alice");
    check sheet.setCell(1, 1, "Engineering");
    check sheet.setCell(2, 0, "Jane");
    // Row 2, col 1 (department) intentionally left blank — a true gap, not "".
    check wb.saveAs(nilMapFile);
    check wb.close();

    map<CellValue>[] data = check parseSheet(nilMapFile, 0, opts);

    test:assertEquals(data.length(), 2, "Should have 2 data rows");
    map<CellValue> jane = data[1];
    test:assertFalse(jane.hasKey("department"),
            "Blank cell's key must be absent when nilAsOptionalField is true");

    check file:remove(nilMapFile);
}

@test:Config {
    groups: ["parseSheet", "projection"]
}
function testParseMapWithNilAsOptionalFieldFalse() returns error? {
    // Test nilAsOptionalField=false for map<CellValue>[] - nil values should be included
    ParseOptions opts = {
        allowDataProjection: {
            nilAsOptionalField: false,
            absentAsNilableType: false
        }
    };

    // Build a fixture with a GENUINELY blank cell (a skipped cell, not an empty
    // string) so the read hits the BLANK → () path. With nilAsOptionalField=false,
    // that column's key must be PRESENT in the map with value ().
    string nilMapFile = TEST_DATA_DIR + "temp_nil_map.xlsx";
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, "name");
    check sheet.setCell(0, 1, "department");
    check sheet.setCell(1, 0, "Alice");
    check sheet.setCell(1, 1, "Engineering");
    check sheet.setCell(2, 0, "Jane");
    // Row 2, col 1 (department) intentionally left blank — a true gap, not "".
    check wb.saveAs(nilMapFile);
    check wb.close();

    map<CellValue>[] data = check parseSheet(nilMapFile, 0, opts);

    test:assertEquals(data.length(), 2, "Should have 2 data rows");
    map<CellValue> jane = data[1];
    test:assertTrue(jane.hasKey("department"), "Blank cell's key must be present in the map");
    test:assertTrue(jane["department"] is (), "Blank cell must bind to () in map<CellValue>");

    check file:remove(nilMapFile);
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
    Workbook wb = check fromFile(TEST_DATA_DIR + "case_headers.xlsx");

    Sheet sheet = check wb.getSheet("Sheet1");
    ParseOptions opts = {
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

    Workbook wb = check fromFile(testFile);
    Sheet sheet = check wb.getSheet(0);

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
function testNegativeHeaderRowErrors() returns error? {
    // A negative headerRowIndex has no valid meaning — the "no headers" sentinel is `()`,
    // not a negative number. Each negative value must surface as a ParseError rather than
    // silently resolving to a negative (impossible) data-start row.
    foreach int badIndex in [-1, -5] {
        ParseOptions opts = {headerRowIndex: badIndex};
        string[][]|Error result = parseSheet(TEST_DATA_DIR + "simple.xlsx", 0, opts);

        test:assertTrue(result is ParseError,
                "headerRowIndex " + badIndex.toString() + " must return ParseError");
        if result is ParseError {
            test:assertTrue(result.message().includes("headerRowIndex"),
                    "Error should name the offending option: " + result.message());
        }
    }
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
    map<CellValue>[] result = check parseSheet(testFile, 0, opts);

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
    // A closed record must DROP sheet columns beyond its declared fields rather than
    // erroring. The fixture has five columns; closed Employee declares only three, so
    // salary/location have no field to land in.
    string testFile = TEST_DATA_DIR + "closed_extra_columns.xlsx";
    string[][] data = [
        ["name", "age", "department", "salary", "location"],
        ["John Doe", "30", "Engineering", "50000", "NYC"],
        ["Jane Smith", "28", "Marketing", "60000", "LA"]
    ];
    check writeSheet(data, testFile);

    // Parsing succeeds despite the two extra columns; only the declared fields populate
    // (the closed record cannot hold the extras).
    Employee[] employees = check parseSheet(testFile);

    test:assertEquals(employees.length(), 2, "Both data rows should parse despite extra columns");
    test:assertEquals(employees[0].name, "John Doe", "First employee name");
    test:assertEquals(employees[0].age, 30, "First employee age");
    test:assertEquals(employees[0].department, "Engineering", "First employee department");
    test:assertEquals(employees[1].name, "Jane Smith", "Second employee name");
    test:assertEquals(employees[1].department, "Marketing", "Second employee department");

    check file:remove(testFile);
}

// =============================================================================
// Natural-type binding for untyped / broad reads (map<CellValue>, rest fields)
// =============================================================================

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseNaturalTypedCellsIntoMap() returns error? {
    // natural_types.xlsx holds genuinely typed cells. Reading into map<CellValue>[]
    // must bind each value to its natural Ballerina type, not collapse to strings.
    map<CellValue>[] data = check parseSheet(TEST_DATA_DIR + "natural_types.xlsx");

    test:assertEquals(data.length(), 1, "Should have 1 data row");
    test:assertEquals(data[0]["intCol"], 42, "Whole number should bind to int");
    test:assertEquals(data[0]["decimalCol"], 3.14d, "Fractional number should bind to decimal");
    test:assertEquals(data[0]["boolCol"], true, "Boolean cell should bind to boolean");
    test:assertEquals(data[0]["dateCol"], "2026-05-28", "Date cell should bind to ISO date string");
    test:assertEquals(data[0]["datetimeCol"], "2026-05-28 14:30:00",
            "Date-time cell should bind to ISO date-time string (time preserved)");
}

@test:Config {
    groups: ["parseSheet", "types"]
}
function testParseNaturalTypedCellsIntoRecord() returns error? {
    // natural_types.xlsx holds genuinely typed cells. Binding them into a strongly
    // typed record must preserve each cell's natural Ballerina type — distinct from
    // the string-cell tests, where the source cells are strings.
    NaturalTypedRow[] data = check parseSheet(TEST_DATA_DIR + "natural_types.xlsx");

    test:assertEquals(data.length(), 1, "Should have 1 data row");
    // intCol is statically `int`, so binding succeeded iff the value asserts equal below.
    test:assertEquals(data[0].intCol, 42, "intCol should bind to int 42");
    test:assertEquals(data[0].decimalCol, 3.14d, "decimalCol should be 3.14");
    test:assertEquals(data[0].boolCol, true, "boolCol should be true");
}

@test:Config {
    groups: ["parseSheet", "openrecord"]
}
function testParseNaturalTypedCellsIntoRestField() returns error? {
    // intCol/boolCol are declared fields; the remaining typed columns fall to the
    // CellValue rest field and must keep their natural types.
    PartialNaturalRow[] rows = check parseSheet(TEST_DATA_DIR + "natural_types.xlsx");

    test:assertEquals(rows.length(), 1);
    test:assertEquals(rows[0].intCol, 42, "Declared int field");
    test:assertEquals(rows[0].boolCol, true, "Declared boolean field");
    test:assertEquals(rows[0]["decimalCol"], 3.14d, "Rest field: fractional → decimal");
    test:assertEquals(rows[0]["dateCol"], "2026-05-28", "Rest field: date → ISO string");
}

// =============================================================================
// Public Data union dispatch — explicit Data target on parse functions
// =============================================================================

@test:Config {groups: ["parseSheet"]}
function testParseSheetWithRowUnionTarget() returns error? {
    // Explicit `Row[]` target — the inferred element is the `Row` union itself.
    // Native dispatch must handle the UNION element tag and fall back to string[][].
    Row[] rows = check parseSheet(TEST_DATA_DIR + "employees.xlsx");
    test:assertTrue(rows is string[][], "Row[] target should fall back to string[][]");
    if rows is string[][] {
        test:assertEquals(rows.length(), 4, "Should have header + 3 data rows");
        test:assertEquals(rows[0], ["name", "age", "department"]);
    }
}

@test:Config {groups: ["parseSheet"]}
function testParseTableWithRowUnionTarget() returns error? {
    // parseTable with explicit Row[] target — same UNION-element dispatch.
    Row[] rows = check parseTable(TEST_DATA_DIR + "tables_test.xlsx", "EmployeeTable");
    test:assertTrue(rows is string[][], "Row[] target on parseTable should fall back to string[][]");
    if rows is string[][] {
        test:assertEquals(rows.length(), 3, "EmployeeTable has 3 data rows");
    }
}

// =============================================================================
// Annotation hardening tests
// =============================================================================
// Annotation values and sheet headers are trimmed on lookup; duplicate
// resolutions (two columns with the same header, or two fields with the same
// @xlsx:Name) fail loud rather than silently collapsing to whichever wins
// the map insertion race.

type EmpWithPaddedAnnotation record {|
    @Name {value: " Department "}
    string department;
    string name;
|};

@test:Config {groups: ["parseSheet", "annotation"]}
function testWhitespaceInXlsxNameAnnotation() returns error? {
    // Write a sheet whose header is exactly "Department" (no spaces).
    string tempFile = getTempFilePath("padded_annot");
    string[][] data = [
        ["Department", "name"],
        ["Engineering", "Alice"]
    ];
    check writeSheet(data, tempFile);

    // Field annotation has " Department " (padded); trim should normalize it
    // to "Department" and the lookup must succeed.
    EmpWithPaddedAnnotation[] employees = check parseSheet(tempFile);
    test:assertEquals(employees.length(), 1);
    test:assertEquals(employees[0].department, "Engineering");
    test:assertEquals(employees[0].name, "Alice");

    check removeTempFile(tempFile);
}

type EmpWithDuplicateAnnotation record {|
    @Name {value: "Label"}
    string a;
    @Name {value: "Label"}
    string b;
|};

@test:Config {groups: ["parseSheet", "annotation"]}
function testDuplicateXlsxNameAnnotationErrors() returns error? {
    string tempFile = getTempFilePath("dup_annot");
    string[][] data = [["Label"], ["value"]];
    check writeSheet(data, tempFile);

    EmpWithDuplicateAnnotation[]|Error result = parseSheet(tempFile);
    test:assertTrue(result is Error,
            "Two fields with the same @xlsx:Name must produce a clear error");

    check removeTempFile(tempFile);
}

type DupHeaderEmp record {|
    string Name;
    string Age;
|};

@test:Config {groups: ["parseSheet", "annotation"]}
function testDuplicateExcelHeaderErrors() returns error? {
    // Build a sheet with two "Name" columns to verify duplicate-header detection.
    string tempFile = getTempFilePath("dup_header");
    string[][] data = [
        ["Name", "Name"],
        ["Alice", "Bob"]
    ];
    check writeSheet(data, tempFile);

    DupHeaderEmp[]|Error result = parseSheet(tempFile);
    test:assertTrue(result is Error,
            "Two sheet columns with the same header must produce a clear error");

    check removeTempFile(tempFile);
}

// =============================================================================
// Boolean coercion safety
// =============================================================================
// CellConverter.parseBoolean used to silently coerce anything outside
// "true"/"yes"/"1" to false. The current behaviour: an explicit false-set
// ("false"/"no"/"0") resolves to false, anything outside either set throws
// TypeConversionException.

type BoolRow record {|
    boolean active;
|};

@test:Config {groups: ["parseSheet", "error"]}
function testInvalidBooleanStringErrors() returns error? {
    string tempFile = getTempFilePath("invalid_boolean");
    string[][] data = [
        ["active"],
        ["maybe"]
    ];
    check writeSheet(data, tempFile);

    BoolRow[]|Error result = parseSheet(tempFile);
    test:assertTrue(result is TypeConversionError,
            "Invalid boolean string 'maybe' must surface as TypeConversionError");
    if result is TypeConversionError {
        test:assertTrue(result.message().includes("to boolean"),
                "Error message should mention the boolean conversion: " + result.message());
    }

    check removeTempFile(tempFile);
}

@test:Config {groups: ["parseSheet"]}
function testExplicitFalseStringsAccepted() returns error? {
    string tempFile = getTempFilePath("explicit_false");
    string[][] data = [
        ["active"],
        ["false"],
        ["no"],
        ["0"]
    ];
    check writeSheet(data, tempFile);

    BoolRow[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 3);
    test:assertEquals(parsed[0].active, false);
    test:assertEquals(parsed[1].active, false);
    test:assertEquals(parsed[2].active, false);

    check removeTempFile(tempFile);
}

// =============================================================================
// Header cells with non-string types should not crash the parser
// =============================================================================
// RecordParsingUtils.buildHeaderMap used to call cell.getStringCellValue()
// directly, which throws IllegalStateException for NUMERIC/BOOLEAN/FORMULA
// cells. The fix routes through CellConverter.convertToStringRaw, which handles
// every cell type uniformly.

type NumericHeaderRow record {|
    @Name {value: "2025"}
    int y1;
    @Name {value: "2026"}
    int y2;
|};

@test:Config {groups: ["parseSheet"]}
function testNumericHeaderCellsAccepted() returns error? {
    // Build a sheet with NUMERIC header cells (years written as integers).
    string tempFile = getTempFilePath("numeric_headers");
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, 2025);
    check sheet.setCell(0, 1, 2026);
    check sheet.setCell(1, 0, 100);
    check sheet.setCell(1, 1, 200);
    check wb.saveAs(tempFile);
    check wb.close();

    NumericHeaderRow[] rows = check parseSheet(tempFile);
    test:assertEquals(rows.length(), 1);
    test:assertEquals(rows[0].y1, 100);
    test:assertEquals(rows[0].y2, 200);

    check removeTempFile(tempFile);
}

// =============================================================================
// Fractional values must not silently truncate when bound to int
// =============================================================================
// CellConverter used to fall back to (long) Double.parseDouble / (long) value
// for int targets, masking precision loss. Both paths now reject fractional
// input via TypeConversionException.

type IntRow record {|
    int n;
|};

@test:Config {groups: ["parseSheet", "error"]}
function testFractionalStringToIntErrors() returns error? {
    string tempFile = getTempFilePath("fractional_int_string");
    string[][] data = [["n"], ["3.7"]];
    check writeSheet(data, tempFile);

    IntRow[]|Error result = parseSheet(tempFile);
    test:assertTrue(result is TypeConversionError,
            "Fractional string '3.7' must surface as TypeConversionError for int target");
    if result is TypeConversionError {
        test:assertTrue(result.message().includes("to int (non-integer value)"),
                "Error message should flag the non-integer value: " + result.message());
    }

    check removeTempFile(tempFile);
}

@test:Config {groups: ["parseSheet", "error"]}
function testFractionalNumericToIntErrors() returns error? {
    string tempFile = getTempFilePath("fractional_int_numeric");
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Data");
    check sheet.setCell(0, 0, "n");
    check sheet.setCell(1, 0, 3.7d);
    check wb.saveAs(tempFile);
    check wb.close();

    IntRow[]|Error result = parseSheet(tempFile);
    test:assertTrue(result is TypeConversionError,
            "Fractional numeric 3.7 must surface as TypeConversionError for int target");
    if result is TypeConversionError {
        test:assertTrue(result.message().includes("to int (non-integer value)"),
                "Error message should flag the non-integer value: " + result.message());
    }

    check removeTempFile(tempFile);
}
