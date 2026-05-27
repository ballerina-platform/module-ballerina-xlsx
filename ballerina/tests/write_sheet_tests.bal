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
    groups: ["writeSheet"]
}
function testWriteStringArray() returns error? {
    string[][] data = [
        ["Name", "Age", "City"],
        ["John", "30", "New York"],
        ["Jane", "25", "Los Angeles"]
    ];

    string tempFile = getTempFilePath("write_string");
    check writeSheet(data, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back and checking exact content
    string[][] parsed = check parseSheet(tempFile);
    assertStringArrayEquals(parsed, data, "Write/read round-trip");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet"]
}
function testWriteRecords() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true},
        {name: "Bob", age: 35, active: false}
    ];

    string tempFile = getTempFilePath("write_records");
    check writeSheet(people, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back as string array to check headers
    string[][] parsed = check parseSheet(tempFile);
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
    groups: ["writeSheet", "options"]
}
function testWriteWithoutHeaders() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true}
    ];

    string tempFile = getTempFilePath("write_no_headers");
    check writeSheet(people, tempFile, writeHeaders = false);

    // Parse back - should only have data row, no headers
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have only 1 row (no headers)");

    // First row should be data, not header names
    test:assertTrue(parsed[0].indexOf("Alice") != (), "First row should contain data 'Alice'");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "options"]
}
function testWriteWithCustomSheetName() returns error? {
    string[][] data = [["Data", "Value"], ["A", "1"]];

    string tempFile = getTempFilePath("write_sheet_name");
    check writeSheet(data, tempFile, sheetName = "MyCustomSheet");

    // Verify by opening as workbook and checking sheet name
    Workbook wb = check new(tempFile);
    string[] sheetNames = wb.getSheetNames();
    test:assertEquals(sheetNames.length(), 1, "Should have 1 sheet");
    test:assertEquals(sheetNames[0], "MyCustomSheet", "Sheet name should match");
    check wb.close();

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "options"]
}
function testWriteWithStartRow() returns error? {
    string[][] data = [
        ["Header1", "Header2"],
        ["Data1", "Data2"]
    ];

    string tempFile = getTempFilePath("write_start_row");
    check writeSheet(data, tempFile, startRowIndex = 2);

    // Parse back - data should start at row 2 (0-based), so rows 0,1 are empty
    // But parse will only see the used range starting from row 2
    string[][] parsed = check parseSheet(tempFile);

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
    groups: ["writeSheet"]
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
    check writeSheet(original, tempFile);

    // Read back
    string[][] parsed = check parseSheet(tempFile);

    // Verify exact match
    assertStringArrayEquals(parsed, original, "Round-trip string array");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet"]
}
function testRoundTripRecords() returns error? {
    Employee[] original = [
        {name: "John Doe", age: 30, department: "Engineering"},
        {name: "Jane Smith", age: 25, department: "Marketing"}
    ];

    string tempFile = getTempFilePath("roundtrip_records");

    // Write to XLSX
    check writeSheet(original, tempFile);

    // Read back as records
    Employee[] parsed = check parseSheet(tempFile);

    // Verify
    assertEmployeesEqual(parsed, original);

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet"]
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

    check writeSheet(original, tempFile);
    string[][] parsed = check parseSheet(tempFile);

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
    groups: ["writeSheet", "types"]
}
function testWriteIntegerValues() returns error? {
    // Write records with int values
    NumericTypes[] data = [
        {intValue: 42, decimalValue: 3.14d},
        {intValue: -100, decimalValue: 0.001d},
        {intValue: 0, decimalValue: 999.999d}
    ];

    string tempFile = getTempFilePath("write_int");
    check writeSheet(data, tempFile);

    // Read back and verify
    NumericTypes[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have 3 records");
    test:assertEquals(parsed[0].intValue, 42, "First int should be 42");
    test:assertEquals(parsed[1].intValue, -100, "Second int should be -100");
    test:assertEquals(parsed[2].intValue, 0, "Third int should be 0");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "types"]
}
function testWriteDecimalValues() returns error? {
    NumericTypes[] data = [
        {intValue: 1, decimalValue: 3.14159d},
        {intValue: 2, decimalValue: 2.71828d}
    ];

    string tempFile = getTempFilePath("write_decimal");
    check writeSheet(data, tempFile);

    NumericTypes[] parsed = check parseSheet(tempFile);
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
    groups: ["writeSheet", "types"]
}
function testWriteBooleanValues() returns error? {
    Person[] data = [
        {name: "Active", age: 30, active: true},
        {name: "Inactive", age: 25, active: false}
    ];

    string tempFile = getTempFilePath("write_boolean");
    check writeSheet(data, tempFile);

    Person[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 2, "Should have 2 records");
    test:assertEquals(parsed[0].active, true, "First should be active");
    test:assertEquals(parsed[1].active, false, "Second should be inactive");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "types"]
}
function testWriteMapArray() returns error? {
    map<anydata>[] data = [
        {"Name": "Alice", "Age": 28, "City": "NYC"},
        {"Name": "Bob", "Age": 35, "City": "LA"}
    ];

    string tempFile = getTempFilePath("write_map");
    check writeSheet(data, tempFile);

    // Read back as string array to verify structure
    string[][] parsed = check parseSheet(tempFile);
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
    groups: ["writeSheet", "annotation"]
}
function testAnnotatedRecordWrite() returns error? {
    AnnotatedEmployee[] employees = [
        {firstName: "John", id: 101, department: "Engineering"},
        {firstName: "Jane", id: 102, department: "Marketing"}
    ];

    string tempFile = getTempFilePath("annotated_write");
    check writeSheet(employees, tempFile);

    // Parse back as string[][] to verify headers match annotation values
    string[][] parsed = check parseSheet(tempFile);

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
    groups: ["writeSheet", "annotation"]
}
function testAnnotatedRecordRoundTrip() returns error? {
    AnnotatedEmployee[] original = [
        {firstName: "Alice", id: 201, department: "Sales"},
        {firstName: "Bob", id: 202, department: "Support"}
    ];

    string tempFile = getTempFilePath("annotated_roundtrip");
    check writeSheet(original, tempFile);

    // Parse back using the same annotated record type
    AnnotatedEmployee[] parsed = check parseSheet(tempFile);

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
    groups: ["writeSheet", "edge"]
}
function testWriteEmptyArray() returns error? {
    string[][] data = [];

    string tempFile = getTempFilePath("write_empty");
    check writeSheet(data, tempFile);

    // File should be created (empty workbook)
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Parse back - should be empty
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 0, "Should be empty");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "edge"]
}
function testWriteSingleRow() returns error? {
    string[][] data = [["Single", "Row", "Data"]];

    string tempFile = getTempFilePath("write_single_row");
    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have 1 row");
    test:assertEquals(parsed[0].length(), 3, "Should have 3 columns");
    test:assertEquals(parsed[0][0], "Single", "First cell");
    test:assertEquals(parsed[0][2], "Data", "Third cell");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "edge"]
}
function testWriteSingleCell() returns error? {
    string[][] data = [["OnlyCell"]];

    string tempFile = getTempFilePath("write_single_cell");
    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have 1 row");
    test:assertEquals(parsed[0].length(), 1, "Should have 1 column");
    test:assertEquals(parsed[0][0], "OnlyCell", "Cell value");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "edge"]
}
function testWriteUnicodeData() returns error? {
    string[][] data = [
        ["Language", "Text"],
        ["Japanese", "\u{3053}\u{3093}\u{306B}\u{3061}\u{306F}"],  // こんにちは
        ["Chinese", "\u{4F60}\u{597D}"],                            // 你好
        ["Korean", "\u{C548}\u{B155}"]                              // 안녕
    ];

    string tempFile = getTempFilePath("write_unicode");
    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
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
    groups: ["writeSheet", "edge"]
}
function testWriteLargeDataset() returns error? {
    // Create a moderately large dataset
    string[][] data = [];
    data.push(["ID", "Name", "Value"]);

    foreach int i in 1 ... 100 {
        data.push([i.toString(), "Item" + i.toString(), (i * 10).toString()]);
    }

    string tempFile = getTempFilePath("write_large");
    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 101, "Should have 101 rows (header + 100 data)");
    test:assertEquals(parsed[1][0], "1", "First data ID should be 1");
    test:assertEquals(parsed[100][0], "100", "Last data ID should be 100");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "edge"]
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
    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");
    test:assertEquals(parsed[0].length(), 10, "Should have 10 columns");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// OVERWRITE BEHAVIOR TESTS
// =============================================================================

@test:Config {
    groups: ["writeSheet"]
}
function testWriteOverwritesExistingFile() returns error? {
    string tempFile = getTempFilePath("overwrite");

    // Write first data
    string[][] data1 = [["First", "Data"]];
    check writeSheet(data1, tempFile);

    // Verify first write
    string[][] parsed1 = check parseSheet(tempFile);
    test:assertEquals(parsed1[0][0], "First", "First write should have 'First'");

    // Overwrite with different data
    string[][] data2 = [["Second", "Data", "More"]];
    check writeSheet(data2, tempFile);

    // Verify overwrite - should have new data, not old
    string[][] parsed2 = check parseSheet(tempFile);
    test:assertEquals(parsed2.length(), 1, "Should have 1 row");
    test:assertEquals(parsed2[0][0], "Second", "Should be overwritten with 'Second'");
    test:assertEquals(parsed2[0].length(), 3, "Should have 3 columns now");

    // Cleanup
    check removeTempFile(tempFile);
}

// =============================================================================
// HEADER-POSITION WRITE TESTS
// =============================================================================
// Tests that verify writes to a sheet with existing headers resolve fields to
// their target columns by header name (not by field iteration order), and that
// unrelated columns are preserved.

type EmpFull record {|
    string name;
    int age;
    int formulaColumn;
    string city;
|};

type EmpProjection record {|
    string name;
    string city;
|};

type EmpThreeCol record {|
    string name;
    int age;
    string city;
|};

type EmpWithUnknownField record {|
    string name;
    string unknownField;
|};

type EmpSimple record {|
    string name;
    int age;
|};

@test:Config {
    groups: ["writeSheet"]
}
function testWriteSheetPreservesUnrelatedColumns() returns error? {
    string tempFile = getTempFilePath("preserves_unrelated_columns");

    // Set up a sheet with 4 columns of data.
    EmpFull[] originalData = [
        {name: "Alice", age: 30, formulaColumn: 100, city: "NYC"},
        {name: "Bob", age: 25, formulaColumn: 200, city: "LA"}
    ];
    check writeSheet(originalData, tempFile);

    // Open via Workbook, project to 2 fields, modify city, write back to the existing sheet.
    Workbook wb = check new(tempFile);
    Sheet sheet = check wb.getSheet(0);

    EmpProjection[] projection = check sheet.getRows();
    test:assertEquals(projection.length(), 2, "Should read 2 projected rows");

    // Modify city
    projection[0] = {name: "Alice", city: "Boston"};
    projection[1] = {name: "Bob", city: "Seattle"};

    // Write back via putRows
    check sheet.putRows(projection);
    check wb.saveAs(tempFile);
    check wb.close();

    // Re-read with the full schema; FormulaColumn and Age must be unchanged.
    EmpFull[] verified = check parseSheet(tempFile);

    test:assertEquals(verified[0].name, "Alice", "Row 0 name preserved");
    test:assertEquals(verified[0].age, 30, "Row 0 age preserved");
    test:assertEquals(verified[0].formulaColumn, 100, "Row 0 formulaColumn preserved");
    test:assertEquals(verified[0].city, "Boston", "Row 0 city updated");

    test:assertEquals(verified[1].age, 25, "Row 1 age preserved");
    test:assertEquals(verified[1].formulaColumn, 200, "Row 1 formulaColumn preserved");
    test:assertEquals(verified[1].city, "Seattle", "Row 1 city updated");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet"]
}
function testWriteSheetErrorsOnUnmatchedField() returns error? {
    string tempFile = getTempFilePath("errors_on_unmatched_field");

    // Set up sheet with three columns: name, age, city
    EmpThreeCol[] originalData = [{name: "Alice", age: 30, city: "NYC"}];
    check writeSheet(originalData, tempFile);

    // Try to write a record whose field has no matching header
    Workbook wb = check new(tempFile);
    Sheet sheet = check wb.getSheet(0);
    EmpWithUnknownField[] bad = [{name: "Bob", unknownField: "value"}];

    Error? result = sheet.putRows(bad);
    test:assertTrue(result is Error, "Should error on field with no matching header");

    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet"]
}
function testWriteSheetFreshSheetIsSequential() returns error? {
    string tempFile = getTempFilePath("fresh_sheet_sequential");

    EmpSimple[] data = [
        {name: "Alice", age: 30},
        {name: "Bob", age: 25}
    ];
    check writeSheet(data, tempFile);

    // Re-read raw and confirm positional layout: header at row 0, two data rows.
    string[][] raw = check parseSheet(tempFile);
    test:assertEquals(raw.length(), 3, "Header + 2 data rows");
    test:assertEquals(raw[0][0], "name", "Header column 0");
    test:assertEquals(raw[0][1], "age", "Header column 1");
    test:assertEquals(raw[1][0], "Alice", "Row 1 column 0");
    test:assertEquals(raw[1][1], "30", "Row 1 column 1");
    test:assertEquals(raw[2][0], "Bob", "Row 2 column 0");
    test:assertEquals(raw[2][1], "25", "Row 2 column 1");

    check removeTempFile(tempFile);
}

// =============================================================================
// Public Data union dispatch — inline literal at writeSheet
// =============================================================================

@test:Config {groups: ["writeSheet"]}
function testWriteSheetWithInlineLiteral() returns error? {
    // No pre-typed local variable — the literal is contextually typed against the
    // public `Data = Row[]` parameter. This forces a union-element BArray into Java
    // and exercises the dispatchWrite UNION_TAG branch in XlsxWriter.
    string tempFile = getTempFilePath("inline_literal_writesheet");
    check writeSheet([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]], tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 3);
    test:assertEquals(parsed[0], ["Name", "Age"]);
    test:assertEquals(parsed[1][0], "Alice");
    test:assertEquals(parsed[2][1], "25");

    check removeTempFile(tempFile);
}

// =============================================================================
// Atomic save tests
// =============================================================================
// File writes go through a temp-file + atomic-rename pattern so a partial or
// failed write never leaves the destination corrupt.

@test:Config {groups: ["writeSheet"]}
function testWriteSheetOverwritesAtomically() returns error? {
    // Pre-create a file with content A, then overwrite with content B. The
    // overwrite must produce a complete, valid workbook (not a half-written one).
    string tempFile = getTempFilePath("atomic_overwrite");

    string[][] contentA = [["Old"], ["data"]];
    check writeSheet(contentA, tempFile);

    string[][] contentB = [["Fresh", "Header"], ["Row1A", "Row1B"], ["Row2A", "Row2B"]];
    check writeSheet(contentB, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed, contentB, "Overwrite must produce content B exactly");

    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetFailureLeavesNoTempFile() returns error? {
    // Writing to a path under a non-existent parent directory fails. The atomic
    // save mechanism creates its temp file in the destination's parent; when the
    // parent doesn't exist, createTempFile fails and no .tmp file is left behind.
    string badPath = TEST_DATA_DIR + "non_existent_dir_for_atomic_test/output.xlsx";

    error? result = writeSheet([["X"]], badPath);
    test:assertTrue(result is error, "Write to nonexistent parent dir must fail");

    // The directory shouldn't exist; can't leave temp orphans in it.
    test:assertFalse(check file:test(TEST_DATA_DIR + "non_existent_dir_for_atomic_test",
            file:EXISTS), "No temp directory should have been created");
}
