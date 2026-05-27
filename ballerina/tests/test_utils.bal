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
import ballerina/jballerina.java;
import ballerina/test;

// =============================================================================
// TEST-ONLY NATIVE EXTERNALS
// =============================================================================
// Bindings to package-private Java helpers used only during fixture generation.
// These are NOT part of the published module — they live in this test file and
// don't compile into the bala. They exist so we can build 1904-epoch fixtures
// programmatically, the way Apache POI / NPOI / openpyxl handle the same need.

function setDate1904Native(Workbook wb, boolean flag) returns Error? = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
} external;

function setFormulaCellNative(Workbook wb, string sheetName, int row, int col, string formula)
        returns Error? = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.xlsx.WorkbookHandle"
} external;

// =============================================================================
// TEST DATA GENERATION
// =============================================================================
// Setup test data files before running tests.
// These files are generated dynamically to ensure consistent test data.

@test:BeforeSuite
function setupTestData() returns error? {
    // Ensure test data directory exists
    if !check file:test(TEST_DATA_DIR, file:EXISTS) {
        check file:createDir(TEST_DATA_DIR);
    }

    // -------------------------------------------------------------------------
    // simple.xlsx - Basic string data
    // -------------------------------------------------------------------------
    string[][] simpleData = [
        ["Name", "Age", "City"],
        ["John", "30", "New York"],
        ["Jane", "25", "Los Angeles"],
        ["Bob", "35", "Chicago"]
    ];
    check writeSheet(simpleData, TEST_DATA_DIR + "simple.xlsx");

    // -------------------------------------------------------------------------
    // employees.xlsx - Typed record data
    // -------------------------------------------------------------------------
    string[][] employeeData = [
        ["name", "age", "department"],
        ["John Doe", "30", "Engineering"],
        ["Jane Smith", "28", "Marketing"],
        ["Bob Johnson", "35", "Sales"]
    ];
    check writeSheet(employeeData, TEST_DATA_DIR + "employees.xlsx");

    // -------------------------------------------------------------------------
    // multi_sheet.xlsx - Multiple sheets for sheet selection tests
    // -------------------------------------------------------------------------
    Workbook wb = check new;

    Sheet sheet1 = check wb.createSheet("Sheet1");
    string[][] sheet1Data = [["A1", "B1"], ["A2", "B2"]];
    check sheet1.putRows(sheet1Data);

    Sheet sheet2 = check wb.createSheet("Sheet2");
    string[][] sheet2Data = [["X1", "Y1"], ["X2", "Y2"]];
    check sheet2.putRows(sheet2Data);

    Sheet sheet3 = check wb.createSheet("Sheet3");
    string[][] sheet3Data = [["P1", "Q1"], ["P2", "Q2"]];
    check sheet3.putRows(sheet3Data);

    check wb.saveAs(TEST_DATA_DIR + "multi_sheet.xlsx");
    check wb.close();

    // -------------------------------------------------------------------------
    // complex_headers.xlsx - Headers not on row 0 (title/logo rows before)
    // -------------------------------------------------------------------------
    string[][] complexHeaderData = [
        ["Company Report", "", ""],           // Row 0: Title
        ["Generated: 2026-01-14", "", ""],    // Row 1: Metadata
        ["Name", "Value", "Status"],          // Row 2: Actual headers
        ["Item1", "100", "Active"],           // Row 3: Data
        ["Item2", "200", "Inactive"]          // Row 4: Data
    ];
    check writeSheet(complexHeaderData, TEST_DATA_DIR + "complex_headers.xlsx");

    // -------------------------------------------------------------------------
    // formulas.xlsx - Cells with real formula cells in column C.
    // Built via the test-only setFormulaCellNative helper because the public
    // write path treats "="-prefixed strings as text (no formula authoring).
    // The cached formula result defaults to "0" because POI doesn't evaluate.
    // -------------------------------------------------------------------------
    Workbook wbFormula = check new;
    Sheet sheetFormula = check wbFormula.createSheet("Formulas");
    check sheetFormula.setCell(0, 0, "A");
    check sheetFormula.setCell(0, 1, "B");
    check sheetFormula.setCell(0, 2, "Sum");
    check sheetFormula.setCell(1, 0, 10);
    check sheetFormula.setCell(1, 1, 20);
    check setFormulaCellNative(wbFormula, "Formulas", 1, 2, "A2+B2");
    check sheetFormula.setCell(2, 0, 15);
    check sheetFormula.setCell(2, 1, 25);
    check setFormulaCellNative(wbFormula, "Formulas", 2, 2, "A3+B3");
    check wbFormula.saveAs(TEST_DATA_DIR + "formulas.xlsx");
    check wbFormula.close();

    // -------------------------------------------------------------------------
    // types_variety.xlsx - Various data types for type conversion tests
    // -------------------------------------------------------------------------
    string[][] typesData = [
        ["text", "number", "amount", "flag"],
        ["Hello", "42", "99.99", "true"],
        ["World", "-10", "0.001", "false"]
    ];
    check writeSheet(typesData, TEST_DATA_DIR + "types_variety.xlsx");

    // -------------------------------------------------------------------------
    // edge_empty_sheet.xlsx - Empty sheet for edge case testing
    // -------------------------------------------------------------------------
    Workbook wbEmpty = check new;
    _ = check wbEmpty.createSheet("EmptySheet");
    check wbEmpty.saveAs(TEST_DATA_DIR + "edge_empty_sheet.xlsx");
    check wbEmpty.close();

    // -------------------------------------------------------------------------
    // edge_single_cell.xlsx - Single cell for edge case testing
    // -------------------------------------------------------------------------
    string[][] singleCellData = [["alone"]];
    check writeSheet(singleCellData, TEST_DATA_DIR + "edge_single_cell.xlsx");

    // -------------------------------------------------------------------------
    // edge_empty_rows.xlsx - Data with empty rows in between
    // -------------------------------------------------------------------------
    Workbook wbGaps = check new;
    Sheet sheetGaps = check wbGaps.createSheet("DataWithGaps");
    // Manually write rows with gaps using putRows at specific positions
    // Row 0: headers, Row 1: data, Row 2: empty, Row 3: data, Row 4: empty, Row 5: data
    string[][] gapsData = [
        ["name", "value"],
        ["First", "100"],
        ["", ""],           // Empty row (will be ghost row)
        ["Second", "200"],
        ["", ""],           // Empty row
        ["Third", "300"]
    ];
    check sheetGaps.putRows(gapsData);
    check wbGaps.saveAs(TEST_DATA_DIR + "edge_empty_rows.xlsx");
    check wbGaps.close();

    // -------------------------------------------------------------------------
    // edge_unicode.xlsx - Unicode characters for internationalization tests
    // -------------------------------------------------------------------------
    string[][] unicodeData = [
        ["Language", "Greeting", "Symbol"],
        ["English", "Hello", "@#$%"],
        ["Japanese", "\u{3053}\u{3093}\u{306B}\u{3061}\u{306F}", "\u{2764}"],     // こんにちは, heart
        ["Arabic", "\u{0645}\u{0631}\u{062D}\u{0628}\u{0627}", "\u{2605}"],       // مرحبا, star
        ["Emoji", "Hello!", "\u{1F600}"]                                          // grinning face
    ];
    check writeSheet(unicodeData, TEST_DATA_DIR + "edge_unicode.xlsx");

    // -------------------------------------------------------------------------
    // annotated.xlsx - Data with headers matching @xlsx:Name annotations
    // -------------------------------------------------------------------------
    string[][] annotatedData = [
        ["First Name", "Employee ID", "Department Name"],
        ["Alice", "101", "Engineering"],
        ["Bob", "102", "Marketing"],
        ["Charlie", "103", "Sales"]
    ];
    check writeSheet(annotatedData, TEST_DATA_DIR + "annotated.xlsx");

    // -------------------------------------------------------------------------
    // numeric_types.xlsx - Numeric values for type conversion
    // -------------------------------------------------------------------------
    string[][] numericData = [
        ["intValue", "decimalValue"],
        ["42", "3.14159"],
        ["-100", "0.001"],
        ["0", "999.999"]
    ];
    check writeSheet(numericData, TEST_DATA_DIR + "numeric_types.xlsx");

    // -------------------------------------------------------------------------
    // nilable_fields.xlsx - Data with missing/empty values
    // -------------------------------------------------------------------------
    string[][] nilableData = [
        ["name", "age", "department"],
        ["John", "30", "Engineering"],
        ["Jane", "", ""],                    // Missing age and department
        ["Bob", "25", ""]                    // Missing department
    ];
    check writeSheet(nilableData, TEST_DATA_DIR + "nilable_fields.xlsx");

    // -------------------------------------------------------------------------
    // case_headers.xlsx - Mixed-case headers for case-insensitive testing
    // Headers: "NAME", "AGE", "Department" (mixed case)
    // Record fields: name, age, department (lowercase)
    // -------------------------------------------------------------------------
    string[][] caseHeadersData = [
        ["NAME", "AGE", "Department"],       // Mixed case headers
        ["John", "30", "Engineering"],
        ["Jane", "25", "Marketing"]
    ];
    check writeSheet(caseHeadersData, TEST_DATA_DIR + "case_headers.xlsx");

    // -------------------------------------------------------------------------
    // dates_1904.xlsx - Workbook with the 1904 date epoch flag set
    // Built programmatically (the public API doesn't expose setDate1904) so we
    // can verify the parser honours `Workbook.isDate1904()` rather than always
    // defaulting to the 1900 epoch. The cell at (0, 0) holds the numeric serial
    // 44708, which is the day count from 1904-01-01 to 2026-05-27. Read via the
    // 1904 epoch this is 2026-05-27; read via the 1900 epoch it would be off
    // by about 4 years and 1 day.
    // -------------------------------------------------------------------------
    Workbook wb1904 = check new;
    check setDate1904Native(wb1904, true);
    Sheet sheet1904 = check wb1904.createSheet("Dates");
    check sheet1904.setCell(0, 0, 44708);
    check wb1904.saveAs(TEST_DATA_DIR + "dates_1904.xlsx");
    check wb1904.close();
}

// Cleanup test data files after running tests
@test:AfterSuite
function cleanupTestData() returns error? {
    string[] filesToRemove = [
        "simple.xlsx",
        "employees.xlsx",
        "multi_sheet.xlsx",
        "complex_headers.xlsx",
        "formulas.xlsx",
        "types_variety.xlsx",
        "edge_empty_sheet.xlsx",
        "edge_single_cell.xlsx",
        "edge_empty_rows.xlsx",
        "edge_unicode.xlsx",
        "annotated.xlsx",
        "numeric_types.xlsx",
        "nilable_fields.xlsx",
        "case_headers.xlsx",
        "extreme_numeric_test.xlsx",
        "dates_1904.xlsx"
    ];

    foreach string fileName in filesToRemove {
        string filePath = TEST_DATA_DIR + fileName;
        if check file:test(filePath, file:EXISTS) {
            check file:remove(filePath);
        }
    }
}

// =============================================================================
// ERROR MESSAGE GENERATORS
// =============================================================================
// These functions generate expected error messages for negative test assertions.

# Generate expected error message for missing sheet.
#
# + sheetName - The sheet name that was not found
# + return - Expected error message string
function generateErrorForMissingSheet(string sheetName) returns string {
    return "Sheet '" + sheetName + "' not found";
}

# Generate expected error message for type conversion failure.
#
# + value - The value that could not be converted
# + targetType - The target type for conversion
# + return - Expected error message string
function generateErrorForTypeConversion(string value, string targetType) returns string {
    return "Cannot convert '" + value + "' to " + targetType;
}

# Generate expected error message for file not found.
#
# + filePath - The file path that was not found
# + return - Expected error message string
function generateErrorForFileNotFound(string filePath) returns string {
    return "Failed to read file: " + filePath;
}

# Generate expected error message for sheet index out of range.
#
# + index - The invalid index
# + maxIndex - The maximum valid index
# + return - Expected error message string
function generateErrorForSheetIndexOutOfRange(int index, int maxIndex) returns string {
    return "Sheet index " + index.toString() + " out of range (0-" + maxIndex.toString() + ")";
}

// =============================================================================
// TEST HELPER FUNCTIONS
// =============================================================================

# Assert that two string arrays are equal (deep comparison).
#
# + actual - The actual string array
# + expected - The expected string array
# + message - Error message prefix
function assertStringArrayEquals(string[][] actual, string[][] expected, string message = "") {
    string prefix = message == "" ? "" : message + ": ";

    test:assertEquals(actual.length(), expected.length(),
        prefix + "Row count mismatch");

    foreach int i in 0 ..< expected.length() {
        test:assertEquals(actual[i].length(), expected[i].length(),
            prefix + "Column count mismatch at row " + i.toString());

        foreach int j in 0 ..< expected[i].length() {
            test:assertEquals(actual[i][j], expected[i][j],
                prefix + "Value mismatch at [" + i.toString() + "][" + j.toString() + "]");
        }
    }
}

# Assert that employee records match expected values.
#
# + actual - The actual employee array
# + expected - The expected employee array
function assertEmployeesEqual(Employee[] actual, Employee[] expected) {
    test:assertEquals(actual.length(), expected.length(), "Employee count mismatch");

    foreach int i in 0 ..< expected.length() {
        test:assertEquals(actual[i].name, expected[i].name,
            "Employee name mismatch at index " + i.toString());
        test:assertEquals(actual[i].age, expected[i].age,
            "Employee age mismatch at index " + i.toString());
        test:assertEquals(actual[i].department, expected[i].department,
            "Employee department mismatch at index " + i.toString());
    }
}

# Create a temporary file path for write tests.
#
# + testName - Name of the test (used in filename)
# + return - Temporary file path
function getTempFilePath(string testName) returns string {
    return TEST_DATA_DIR + "temp_" + testName + ".xlsx";
}

# Remove a temporary test file if it exists.
#
# + filePath - Path to the file to remove
# + return - Error if removal fails
function removeTempFile(string filePath) returns error? {
    if check file:test(filePath, file:EXISTS) {
        check file:remove(filePath);
    }
}
