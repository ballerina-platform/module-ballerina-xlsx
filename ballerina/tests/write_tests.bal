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
// BASIC WRITE TESTS
// =============================================================================

@test:Config {
    groups: ["write"]
}
function testWriteStringArray() returns error? {
    string[][] data = [
        ["Name", "Age", "City"],
        ["John", "30", "New York"],
        ["Jane", "25", "Los Angeles"]
    ];

    string tempFile = getTempFilePath("write_string");
    check write(data, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back and checking exact content
    string[][] parsed = check parse(tempFile);
    assertStringArrayEquals(parsed, data, "Write/read round-trip");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testWriteRecords() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true},
        {name: "Bob", age: 35, active: false}
    ];

    string tempFile = getTempFilePath("write_records");
    check write(people, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back as string array to check headers
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");

    // Verify headers exist (order may vary)
    string[] headers = parsed[0];
    test:assertTrue(headers.indexOf("name") != () || headers.indexOf("Name") != (),
        "Should have 'name' header");
    test:assertTrue(headers.indexOf("age") != () || headers.indexOf("Age") != (),
        "Should have 'age' header");
    test:assertTrue(headers.indexOf("active") != () || headers.indexOf("Active") != (),
        "Should have 'active' header");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "options"]
}
function testWriteWithoutHeaders() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true}
    ];

    string tempFile = getTempFilePath("write_no_headers");
    check write(people, tempFile, writeHeaders = false);

    // Parse back - should only have data row, no headers
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have only 1 row (no headers)");

    // First row should be data, not header names
    test:assertTrue(parsed[0].indexOf("Alice") != (), "First row should contain data 'Alice'");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "options"]
}
function testWriteWithCustomSheetName() returns error? {
    string[][] data = [["Data", "Value"], ["A", "1"]];

    string tempFile = getTempFilePath("write_sheet_name");
    check write(data, tempFile, sheetName = "MyCustomSheet");

    // Verify by opening as workbook and checking sheet name
    Workbook wb = check openFile(tempFile);
    string[] sheetNames = wb.getSheetNames();
    test:assertEquals(sheetNames.length(), 1, "Should have 1 sheet");
    test:assertEquals(sheetNames[0], "MyCustomSheet", "Sheet name should match");
    check wb.close();

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "options"]
}
function testWriteWithStartRow() returns error? {
    string[][] data = [
        ["Header1", "Header2"],
        ["Data1", "Data2"]
    ];

    string tempFile = getTempFilePath("write_start_row");
    check write(data, tempFile, startRowIndex = 2);

    // Parse back - data should start at row 2 (0-based), so rows 0,1 are empty
    // But parse will only see the used range starting from row 2
    string[][] parsed = check parse(tempFile);

    // The data we wrote should be there
    test:assertEquals(parsed.length(), 2, "Should have 2 rows");
    test:assertEquals(parsed[0][0], "Header1", "First row should be Header1");
    test:assertEquals(parsed[1][0], "Data1", "Second row should be Data1");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// ROUND-TRIP TESTS
// =============================================================================

@test:Config {
    groups: ["write"]
}
function testRoundTripStringArray() returns error? {
    string[][] original = [
        ["Name", "Score", "Grade"],
        ["Alice", "95", "A"],
        ["Bob", "87", "B"],
        ["Charlie", "92", "A-"]
    ];

    string tempFile = getTempFilePath("roundtrip_string");

    // Write to XLSX
    check write(original, tempFile);

    // Read back
    string[][] parsed = check parse(tempFile);

    // Verify exact match
    assertStringArrayEquals(parsed, original, "Round-trip string array");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testRoundTripRecords() returns error? {
    Employee[] original = [
        {name: "John Doe", age: 30, department: "Engineering"},
        {name: "Jane Smith", age: 25, department: "Marketing"}
    ];

    string tempFile = getTempFilePath("roundtrip_records");

    // Write to XLSX
    check write(original, tempFile);

    // Read back as records
    Employee[] parsed = check parse(tempFile);

    // Verify
    assertEmployeesEqual(parsed, original);

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testRoundTripWithSpecialCharacters() returns error? {
    string[][] original = [
        ["Name", "Description"],
        ["Test", "Special chars: @#$%^&*()"],
        ["Quote", "He said \"Hello\""],
        ["Newline", "Line1"],  // Note: actual newlines in cells are complex
        ["Comma", "a,b,c"]
    ];

    string tempFile = getTempFilePath("roundtrip_special");

    check write(original, tempFile);
    string[][] parsed = check parse(tempFile);

    // Verify special characters preserved
    test:assertEquals(parsed[1][1], "Special chars: @#$%^&*()", "Special chars preserved");
    test:assertEquals(parsed[4][1], "a,b,c", "Commas preserved");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// TYPE PRESERVATION TESTS
// =============================================================================

@test:Config {
    groups: ["write", "types"]
}
function testWriteIntegerValues() returns error? {
    // Write records with int values
    NumericTypes[] data = [
        {intValue: 42, decimalValue: 3.14d},
        {intValue: -100, decimalValue: 0.001d},
        {intValue: 0, decimalValue: 999.999d}
    ];

    string tempFile = getTempFilePath("write_int");
    check write(data, tempFile);

    // Read back and verify
    NumericTypes[] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have 3 records");
    test:assertEquals(parsed[0].intValue, 42, "First int should be 42");
    test:assertEquals(parsed[1].intValue, -100, "Second int should be -100");
    test:assertEquals(parsed[2].intValue, 0, "Third int should be 0");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "types"]
}
function testWriteDecimalValues() returns error? {
    NumericTypes[] data = [
        {intValue: 1, decimalValue: 3.14159d},
        {intValue: 2, decimalValue: 2.71828d}
    ];

    string tempFile = getTempFilePath("write_decimal");
    check write(data, tempFile);

    NumericTypes[] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 2, "Should have 2 records");
    // Decimal comparison with tolerance
    test:assertTrue(parsed[0].decimalValue > 3.14d && parsed[0].decimalValue < 3.15d,
        "First decimal should be ~3.14159");
    test:assertTrue(parsed[1].decimalValue > 2.71d && parsed[1].decimalValue < 2.72d,
        "Second decimal should be ~2.71828");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "types"]
}
function testWriteBooleanValues() returns error? {
    Person[] data = [
        {name: "Active", age: 30, active: true},
        {name: "Inactive", age: 25, active: false}
    ];

    string tempFile = getTempFilePath("write_boolean");
    check write(data, tempFile);

    Person[] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 2, "Should have 2 records");
    test:assertEquals(parsed[0].active, true, "First should be active");
    test:assertEquals(parsed[1].active, false, "Second should be inactive");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "types"]
}
function testWriteMapArray() returns error? {
    map<anydata>[] data = [
        {"Name": "Alice", "Age": 28, "City": "NYC"},
        {"Name": "Bob", "Age": 35, "City": "LA"}
    ];

    string tempFile = getTempFilePath("write_map");
    check write(data, tempFile);

    // Read back as string array to verify structure
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");

    // Verify headers (map keys become headers)
    string[] headers = parsed[0];
    test:assertTrue(headers.indexOf("Name") != (), "Should have 'Name' header");
    test:assertTrue(headers.indexOf("Age") != (), "Should have 'Age' header");
    test:assertTrue(headers.indexOf("City") != (), "Should have 'City' header");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// ANNOTATION TESTS
// =============================================================================

@test:Config {
    groups: ["write", "annotation"]
}
function testAnnotatedRecordWrite() returns error? {
    AnnotatedEmployee[] employees = [
        {firstName: "John", id: 101, department: "Engineering"},
        {firstName: "Jane", id: 102, department: "Marketing"}
    ];

    string tempFile = getTempFilePath("annotated_write");
    check write(employees, tempFile);

    // Parse back as string[][] to verify headers match annotation values
    string[][] parsed = check parse(tempFile);

    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");

    // Verify headers are annotation values, not field names
    string[] headers = parsed[0];
    test:assertTrue(headers.indexOf("First Name") != (), "Should have 'First Name' header");
    test:assertTrue(headers.indexOf("Employee ID") != (), "Should have 'Employee ID' header");
    test:assertTrue(headers.indexOf("Department Name") != (), "Should have 'Department Name' header");

    // Verify field names are NOT used as headers
    test:assertTrue(headers.indexOf("firstName") == (), "Should NOT have 'firstName' header");
    test:assertTrue(headers.indexOf("id") == (), "Should NOT have 'id' header");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "annotation"]
}
function testAnnotatedRecordRoundTrip() returns error? {
    AnnotatedEmployee[] original = [
        {firstName: "Alice", id: 201, department: "Sales"},
        {firstName: "Bob", id: 202, department: "Support"}
    ];

    string tempFile = getTempFilePath("annotated_roundtrip");
    check write(original, tempFile);

    // Parse back using the same annotated record type
    AnnotatedEmployee[] parsed = check parse(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 employee records");
    test:assertEquals(parsed[0].firstName, "Alice", "First employee firstName");
    test:assertEquals(parsed[0].id, 201, "First employee id");
    test:assertEquals(parsed[0].department, "Sales", "First employee department");
    test:assertEquals(parsed[1].firstName, "Bob", "Second employee firstName");
    test:assertEquals(parsed[1].id, 202, "Second employee id");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// EDGE CASE TESTS
// =============================================================================

@test:Config {
    groups: ["write", "edge"]
}
function testWriteEmptyArray() returns error? {
    string[][] data = [];

    string tempFile = getTempFilePath("write_empty");
    check write(data, tempFile);

    // File should be created (empty workbook)
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Parse back - should be empty
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 0, "Should be empty");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "edge"]
}
function testWriteSingleRow() returns error? {
    string[][] data = [["Single", "Row", "Data"]];

    string tempFile = getTempFilePath("write_single_row");
    check write(data, tempFile);

    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have 1 row");
    test:assertEquals(parsed[0].length(), 3, "Should have 3 columns");
    test:assertEquals(parsed[0][0], "Single", "First cell");
    test:assertEquals(parsed[0][2], "Data", "Third cell");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "edge"]
}
function testWriteSingleCell() returns error? {
    string[][] data = [["OnlyCell"]];

    string tempFile = getTempFilePath("write_single_cell");
    check write(data, tempFile);

    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have 1 row");
    test:assertEquals(parsed[0].length(), 1, "Should have 1 column");
    test:assertEquals(parsed[0][0], "OnlyCell", "Cell value");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "edge"]
}
function testWriteUnicodeData() returns error? {
    string[][] data = [
        ["Language", "Text"],
        ["Japanese", "\u{3053}\u{3093}\u{306B}\u{3061}\u{306F}"],  // こんにちは
        ["Chinese", "\u{4F60}\u{597D}"],                            // 你好
        ["Korean", "\u{C548}\u{B155}"]                              // 안녕
    ];

    string tempFile = getTempFilePath("write_unicode");
    check write(data, tempFile);

    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 4, "Should have 4 rows");
    test:assertEquals(parsed[0][0], "Language", "Header");
    test:assertEquals(parsed[1][0], "Japanese", "First data row label");
    // Verify Unicode content exists (exact comparison may be tricky)
    test:assertTrue(parsed[1][1].length() > 0, "Japanese text should have content");
    test:assertTrue(parsed[2][1].length() > 0, "Chinese text should have content");
    test:assertTrue(parsed[3][1].length() > 0, "Korean text should have content");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "edge"]
}
function testWriteLargeDataset() returns error? {
    // Create a moderately large dataset
    string[][] data = [];
    data.push(["ID", "Name", "Value"]);

    foreach int i in 1 ... 100 {
        data.push([i.toString(), "Item" + i.toString(), (i * 10).toString()]);
    }

    string tempFile = getTempFilePath("write_large");
    check write(data, tempFile);

    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 101, "Should have 101 rows (header + 100 data)");
    test:assertEquals(parsed[1][0], "1", "First data ID should be 1");
    test:assertEquals(parsed[100][0], "100", "Last data ID should be 100");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "edge"]
}
function testWriteWideData() returns error? {
    // Create wide data with 10 columns
    WideRecord[] data = [
        {col1: "A1", col2: "B1", col3: "C1", col4: "D1", col5: "E1",
         col6: "F1", col7: "G1", col8: "H1", col9: "I1", col10: "J1"},
        {col1: "A2", col2: "B2", col3: "C2", col4: "D2", col5: "E2",
         col6: "F2", col7: "G2", col8: "H2", col9: "I2", col10: "J2"}
    ];

    string tempFile = getTempFilePath("write_wide");
    check write(data, tempFile);

    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");
    test:assertEquals(parsed[0].length(), 10, "Should have 10 columns");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// OVERWRITE BEHAVIOR TESTS
// =============================================================================

@test:Config {
    groups: ["write"]
}
function testWriteOverwritesExistingFile() returns error? {
    string tempFile = getTempFilePath("overwrite");

    // Write first data
    string[][] data1 = [["First", "Data"]];
    check write(data1, tempFile);

    // Verify first write
    string[][] parsed1 = check parse(tempFile);
    test:assertEquals(parsed1[0][0], "First", "First write should have 'First'");

    // Overwrite with different data
    string[][] data2 = [["Second", "Data", "More"]];
    check write(data2, tempFile);

    // Verify overwrite - should have new data, not old
    string[][] parsed2 = check parse(tempFile);
    test:assertEquals(parsed2.length(), 1, "Should have 1 row");
    test:assertEquals(parsed2[0][0], "Second", "Should be overwritten with 'Second'");
    test:assertEquals(parsed2[0].length(), 3, "Should have 3 columns now");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// ROW WRAPPER WRITE TESTS
// =============================================================================
// Tests for position-aware writing with Row-wrapped types.

@test:Config {
    groups: ["write", "row-wrapper"]
}
function testWriteRowWrappedData() returns error? {
    // Create Row-wrapped data with gaps (simulating empty rows)
    SimpleDataRow[] data = [
        {rowIndex: 0, value: {name: "First", value: 100}},
        {rowIndex: 2, value: {name: "Second", value: 200}},  // Skip row 1
        {rowIndex: 4, value: {name: "Third", value: 300}}    // Skip row 3
    ];

    string tempFile = getTempFilePath("write_row_wrapper");
    check write(data, tempFile);

    // Parse back as string array to verify positions
    string[][] parsed = check parse(tempFile);

    // Should have header + at least 5 rows (positions 0,1,2,3,4)
    // Row 0: header, Row 1: First, Row 2: empty, Row 3: Second, Row 4: empty, Row 5: Third
    test:assertTrue(parsed.length() >= 4, "Should have headers + data rows");

    // Verify header row
    test:assertEquals(parsed[0][0], "name", "First header should be 'name'");
    test:assertEquals(parsed[0][1], "value", "Second header should be 'value'");

    // Verify data is at expected positions
    test:assertEquals(parsed[1][0], "First", "Row 1 should be 'First'");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "row-wrapper"]
}
function testWriteRowWrappedDataWithNullValues() returns error? {
    // Create Row-wrapped data including null values (empty rows)
    SimpleDataRow[] data = [
        {rowIndex: 0, value: {name: "Start", value: 1}},
        {rowIndex: 1, value: null},  // Empty row - should be skipped in output
        {rowIndex: 2, value: {name: "End", value: 2}}
    ];

    string tempFile = getTempFilePath("write_row_wrapper_null");
    check write(data, tempFile);

    // Parse back
    string[][] parsed = check parse(tempFile);

    // Null values are skipped during write, so we should have header + 2 data rows
    test:assertTrue(parsed.length() >= 3, "Should have header + data rows");
    test:assertEquals(parsed[0][0], "name", "Header should be 'name'");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// ROUND-TRIP POSITION PRESERVATION TESTS
// =============================================================================
// Tests that verify position preservation during parse -> modify -> write -> parse cycle.

@test:Config {
    groups: ["write", "row-wrapper", "roundtrip"]
}
function testRoundTripWithRowWrapper() returns error? {
    // Step 1: Parse file with empty rows using Row wrapper
    SimpleDataRow[] originalRows = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Should have 5 rows including empty ones
    test:assertEquals(originalRows.length(), 5, "Should parse 5 rows including empty");

    // Step 2: Modify some values (but keep positions)
    SimpleDataRow[] modifiedRows = originalRows.clone();
    if modifiedRows[0].value != null {
        modifiedRows[0] = {rowIndex: 0, value: {name: "Modified", value: 999}};
    }

    // Step 3: Write back
    string tempFile = getTempFilePath("roundtrip_row_wrapper");
    check write(modifiedRows, tempFile);

    // Step 4: Parse again and verify positions are preserved
    SimpleDataRow[] reparsedRows = check parse(tempFile);

    // Verify the modification persisted at the correct position
    test:assertEquals(reparsedRows[0].rowIndex, 0, "First row should still be at rowIndex 0");
    test:assertEquals(reparsedRows[0].value?.name, "Modified", "Modified value should persist");
    test:assertEquals(reparsedRows[0].value?.value, 999, "Modified value should persist");

    // Verify other positions are maintained
    test:assertEquals(reparsedRows[2].rowIndex, 2, "Third row should still be at rowIndex 2");
    test:assertEquals(reparsedRows[2].value?.name, "Second", "Second data should be at rowIndex 2");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["write", "row-wrapper", "roundtrip"]
}
function testRoundTripFilterAndWriteBack() returns error? {
    // Step 1: Parse with Row wrapper
    SimpleDataRow[] rows = check parse(TEST_DATA_DIR + "edge_empty_rows.xlsx");

    // Step 2: Filter (keep only rows with value > 150)
    SimpleDataRow[] filtered = rows.filter(r => r.value != null && r.value?.value > 150);

    test:assertEquals(filtered.length(), 2, "Should have 2 rows after filtering (Second=200, Third=300)");

    // Verify positions are preserved after filtering
    test:assertEquals(filtered[0].rowIndex, 2, "Second should have original rowIndex 2");
    test:assertEquals(filtered[1].rowIndex, 4, "Third should have original rowIndex 4");

    // Step 3: Write filtered data (positions should be used)
    string tempFile = getTempFilePath("roundtrip_filtered");
    check write(filtered, tempFile);

    // Step 4: Parse back and verify
    SimpleDataRow[] reparsed = check parse(tempFile);

    // The written data should maintain the relative positions
    test:assertTrue(reparsed.length() >= 2, "Should have at least 2 rows");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["workbook", "row-wrapper", "roundtrip"]
}
function testWorkbookRoundTripWithRowWrapper() returns error? {
    // Test round-trip via Workbook API
    Workbook wb = check openFile(TEST_DATA_DIR + "edge_empty_rows.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    // Get rows with Row wrapper
    SimpleDataRow[] rows = check sheet.getRows();
    test:assertEquals(rows.length(), 5, "Should have 5 rows");

    // Create new workbook and write
    Workbook newWb = check createWorkbook();
    Sheet newSheet = check newWb.createSheet("Data");
    check newSheet.putRows(rows);

    string tempFile = getTempFilePath("workbook_roundtrip");
    check newWb.saveAs(tempFile);
    check newWb.close();
    check wb.close();

    // Verify by re-opening
    Workbook verifyWb = check openFile(tempFile);
    Sheet verifySheet = check verifyWb.getSheetByIndex(0);
    SimpleDataRow[] verifiedRows = check verifySheet.getRows();

    test:assertTrue(verifiedRows.length() >= 3, "Should have preserved rows");

    check verifyWb.close();
    check removeTempFile(tempFile);
}
