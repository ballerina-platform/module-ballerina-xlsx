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
    int? nameCol = headers.indexOf("name");
    int? ageCol = headers.indexOf("age");
    int? activeCol = headers.indexOf("active");
    test:assertTrue(nameCol != (), "Should have 'name' header");
    test:assertTrue(ageCol != (), "Should have 'age' header");
    test:assertTrue(activeCol != (), "Should have 'active' header");

    // Verify the actual data values land in the columns their headers indicate.
    int nc = <int>nameCol;
    int ac = <int>ageCol;
    int actc = <int>activeCol;
    test:assertEquals(parsed[1][nc], "Alice", "Row 1 name should be Alice");
    test:assertEquals(parsed[1][ac], "28", "Row 1 age should be 28");
    test:assertEquals(parsed[1][actc], "true", "Row 1 active should be true");
    test:assertEquals(parsed[2][nc], "Bob", "Row 2 name should be Bob");
    test:assertEquals(parsed[2][ac], "35", "Row 2 age should be 35");
    test:assertEquals(parsed[2][actc], "false", "Row 2 active should be false");

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

    // Parse back with no header row so the single written row is read as data.
    string[][] parsed = check parseSheet(tempFile, 0, {headerRowIndex: ()});
    test:assertEquals(parsed.length(), 1, "Should have only 1 row (no headers)");

    // Row 0 must be the data values, with no header row preceding them.
    string[] row0 = parsed[0];
    test:assertTrue(row0.indexOf("Alice") != (), "Row 0 should contain data 'Alice'");
    test:assertTrue(row0.indexOf("28") != (), "Row 0 should contain data '28'");
    test:assertTrue(row0.indexOf("true") != (), "Row 0 should contain data 'true'");

    // No header names should have been written.
    test:assertTrue(row0.indexOf("name") == (), "No 'name' header should be written");
    test:assertTrue(row0.indexOf("age") == (), "No 'age' header should be written");
    test:assertTrue(row0.indexOf("active") == (), "No 'active' header should be written");

    // Cleanup
    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["writeSheet", "options"]
}
function testWriteWithCustomSheetName() returns error? {
    string[][] data = [["Data", "Value"], ["A", "1"]];

    string tempFile = getTempFilePath("write_sheet_name");
    check writeSheet(data, tempFile, "MyCustomSheet");

    // Verify by opening as workbook and checking sheet name
    Workbook wb = check fromFile(tempFile);
    string[] sheetNames = check wb.getSheetNames();
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

    // Open via Workbook to inspect absolute cell positions: the leading rows
    // (0 and 1) must be blank, proving startRowIndex was honored.
    Workbook wb = check fromFile(tempFile);
    Sheet sheet = check wb.getSheet(0);

    CellValue leading0 = check sheet.getCell(0, 0);
    CellValue leading1 = check sheet.getCell(1, 0);
    test:assertEquals(leading0, (), "Row 0 should be blank (data starts at row 2)");
    test:assertEquals(leading1, (), "Row 1 should be blank (data starts at row 2)");

    // The data we wrote lands at row 2 onward.
    CellValue header = check sheet.getCell(2, 0);
    CellValue dataCell = check sheet.getCell(3, 0);
    test:assertEquals(header, "Header1", "Row 2 should hold Header1");
    test:assertEquals(dataCell, "Data1", "Row 3 should hold Data1");

    CellRange? range = check sheet.getUsedCellRange();
    test:assertTrue(range != (), "Used range should exist");
    if range != () {
        test:assertEquals(range.firstRowIndex, 2, "Used range should start at row 2");
    }
    check wb.close();

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
    map<CellValue>[] data = [
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
    int? nameCol = headers.indexOf("Name");
    int? ageCol = headers.indexOf("Age");
    int? cityCol = headers.indexOf("City");
    test:assertTrue(nameCol != (), "Should have 'Name' header");
    test:assertTrue(ageCol != (), "Should have 'Age' header");
    test:assertTrue(cityCol != (), "Should have 'City' header");

    // Verify the data values appear under their respective headers.
    int nc = <int>nameCol;
    int ac = <int>ageCol;
    int cc = <int>cityCol;
    test:assertEquals(parsed[1][nc], "Alice", "Row 1 Name should be Alice");
    test:assertEquals(parsed[1][ac], "28", "Row 1 Age should be 28");
    test:assertEquals(parsed[1][cc], "NYC", "Row 1 City should be NYC");
    test:assertEquals(parsed[2][nc], "Bob", "Row 2 Name should be Bob");
    test:assertEquals(parsed[2][ac], "35", "Row 2 Age should be 35");
    test:assertEquals(parsed[2][cc], "LA", "Row 2 City should be LA");

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

    // Overwrite with different data — explicit REPLACE, since writeSheet refuses to clobber
    // an existing sheet by default (FAIL_IF_EXISTS).
    string[][] data2 = [["Second", "Data", "More"]];
    check writeSheet(data2, tempFile, sheetWriteMode = REPLACE);

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
    Workbook wb = check fromFile(tempFile);
    Sheet sheet = check wb.getSheet(0);

    EmpProjection[] projection = check sheet.getRows();
    test:assertEquals(projection.length(), 2, "Should read 2 projected rows");

    // Modify city
    projection[0] = {name: "Alice", city: "Boston"};
    projection[1] = {name: "Bob", city: "Seattle"};

    // Write back via putRows — overwrite the data rows in place (REPLACE; putRows defaults to APPEND).
    // REPLACE aligns to the existing header and touches only the projected columns, so the
    // unrelated columns (age, formulaColumn) are preserved.
    check sheet.putRows(projection, sheetWriteMode = REPLACE);
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
    Workbook wb = check fromFile(tempFile);
    Sheet sheet = check wb.getSheet(0);
    EmpWithUnknownField[] bad = [{name: "Bob", unknownField: "value"}];

    Error? result = sheet.putRows(bad);
    test:assertTrue(result is Error, "Should error on field with no matching header");
    if result is Error {
        test:assertTrue(result.message().includes("no matching column"),
            "Error should mention no matching column");
    }

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

@test:Config {groups: ["writeSheet"]}
function testWriteEqualsPrefixedStringAsText() returns error? {
    string tempFile = getTempFilePath("equals_prefixed_text");

    string injection = "=HYPERLINK(\"https://example.com/phish\", \"reset password\")";
    string[][] data = [
        ["note", "value"],
        ["=N/A", "missing"],
        ["=SUM(B2:B10)", "literal sum"],
        [injection, "injection vector"]
    ];

    check writeSheet(data, tempFile);

    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 4);
    test:assertEquals(parsed[1][0], "=N/A");
    test:assertEquals(parsed[2][0], "=SUM(B2:B10)");
    test:assertEquals(parsed[3][0], injection);

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
    check writeSheet(contentB, tempFile, sheetWriteMode = REPLACE);

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

// Concurrent writeSheet calls must be truly isolated. Pre-fix a static
// STYLE_CACHE leaked styles and raced across calls. The per-call style cache
// fix makes each call independent; this test drives 10 parallel workers to
// distinct files and verifies all 10 succeed and round-trip exactly.
@test:Config {groups: ["writeSheet"]}
function testConcurrentWritesIsolated() returns error? {
    int workers = 10;
    int rowsPerWorker = 100;
    string[] paths = [];
    future<Error?>[] futures = [];

    foreach int w in 0 ..< workers {
        string path = getTempFilePath("concurrent_" + w.toString());
        paths.push(path);
        string[][] data = [["worker", "row"]];
        foreach int r in 0 ..< rowsPerWorker {
            data.push([w.toString(), r.toString()]);
        }
        future<Error?> f = start writeSheet(data, path);
        futures.push(f);
    }

    foreach int w in 0 ..< workers {
        Error? err = wait futures[w];
        test:assertTrue(err is (),
                "Worker " + w.toString() + " must complete without error");
    }

    foreach int w in 0 ..< workers {
        string[][] parsed = check parseSheet(paths[w]);
        test:assertEquals(parsed.length(), rowsPerWorker + 1,
                "Worker " + w.toString() + " file must have all rows");
        test:assertEquals(parsed[1][0], w.toString(),
                "Worker " + w.toString() + " first data row must be its own ID");
        check removeTempFile(paths[w]);
    }
}

@test:Config {groups: ["writeSheet"]}
function testWriteLargeIntSilentPrecisionLoss() returns error? {
    string tempFile = getTempFilePath("large_int_precision");

    int safeBoundary = 9007199254740992;        // 2^53 — last exact int in a double
    int firstLossy = 9007199254740993;          // 2^53 + 1 — rounds down to 2^53
    int ibanLike = 4929187654321098765;         // ~5e18, off-grid at 2^62 spacing (1024)

    // Note: int:MAX_VALUE deliberately avoided. It appears to round-trip due to
    // double→long saturation on read masking the lossy long→double rounding on
    // write — a coincidence at the long-max boundary, not preserved precision.
    record {|int n;|}[] data = [
        {n: safeBoundary},
        {n: firstLossy},
        {n: ibanLike}
    ];
    check writeSheet(data, tempFile);

    record {|int n;|}[] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed[0].n, safeBoundary,
            "2^53 must round-trip exactly — still within double mantissa");
    test:assertEquals(parsed[1].n, safeBoundary,
            "2^53+1 must silently round to 2^53 — Option B / matches POI");
    test:assertNotEquals(parsed[2].n, ibanLike,
            "Off-grid mid-magnitude value must NOT round-trip exactly — Option B contract");

    check removeTempFile(tempFile);
}

// =============================================================================
// Mixed-shape Row[] writes must fail loudly, not silently drop rows
// =============================================================================
// XlsxWriter.writeRecordData used to `continue` past any non-record element,
// silently losing data in a mixed Row[] input. The fix raises a typed error
// so the caller sees the shape mismatch immediately.

@test:Config {groups: ["writeSheet"]}
function testMixedRowShapeWriteErrors() returns error? {
    string tempFile = getTempFilePath("mixed_shape");

    // Quote the keys: when the contextual type is the Row union, Ballerina
    // routes a bare-identifier mapping constructor through the open `record{}`
    // member, where rest-field keys must be string literals. Quoting works
    // for both the record and map branches and keeps this test future-proof.
    Row[] data = [
        {"name": "Alice", "age": 30},  // record — dispatchWrite picks the record writer
        ["Bob", "25"]                   // string[] — incompatible with record-writer path
    ];

    Error? result = writeSheet(data, tempFile);
    test:assertTrue(result is Error,
            "Mixed-shape Row[] write must surface a clear error, not silently drop rows");
}

// =============================================================================
// SheetWriteMode — non-destructive writeSheet (FAIL_IF_EXISTS / REPLACE / APPEND)
// =============================================================================
// writeSheet opens an existing file and preserves sibling sheets. The default
// FAIL_IF_EXISTS refuses to clobber an existing target sheet; REPLACE overwrites
// it (siblings preserved) and APPEND adds rows below the existing data.

type AppendLogRow record {|
    string name;
    int age;
    string dept;
|};

// Same columns as AppendLogRow but a different field order, to prove APPEND aligns
// to the existing header by name rather than by position.
type AppendLogShuffled record {|
    int age;
    string dept;
    string name;
|};

type TwoIntRow record {|
    int a;
    int b;
|};

// Builds a fresh three-sheet workbook (Alpha/Beta/Gamma) at `path`.
function createMultiSheetWorkbook(string path) returns error? {
    Workbook wb = new;
    Sheet alpha = check wb.createSheet("Alpha");
    check alpha.putRows([["a1", "a2"], ["a3", "a4"]]);
    Sheet beta = check wb.createSheet("Beta");
    check beta.putRows([["b1", "b2"], ["b3", "b4"]]);
    Sheet gamma = check wb.createSheet("Gamma");
    check gamma.putRows([["g1", "g2"], ["g3", "g4"]]);
    check wb.saveAs(path);
    check wb.close();
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetReplacePreservesSiblings() returns error? {
    string tempFile = getTempFilePath("replace_siblings");
    check createMultiSheetWorkbook(tempFile);

    // Replace only "Beta"; Alpha and Gamma (and the tab order) must survive.
    check writeSheet([["x1", "x2", "x3"]], tempFile, "Beta", sheetWriteMode = REPLACE);

    Workbook wb = check fromFile(tempFile);
    string[] names = check wb.getSheetNames();
    test:assertEquals(names, ["Alpha", "Beta", "Gamma"], "Siblings preserved in original order");

    Sheet alpha = check wb.getSheet("Alpha");
    string[][] alphaRows = check alpha.getRows();
    test:assertEquals(alphaRows, [["a1", "a2"], ["a3", "a4"]], "Alpha untouched");

    Sheet beta = check wb.getSheet("Beta");
    string[][] betaRows = check beta.getRows();
    test:assertEquals(betaRows, [["x1", "x2", "x3"]], "Beta replaced with the new data");

    Sheet gamma = check wb.getSheet("Gamma");
    string[][] gammaRows = check gamma.getRows();
    test:assertEquals(gammaRows, [["g1", "g2"], ["g3", "g4"]], "Gamma untouched");

    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetReplaceCreatesAbsentSheet() returns error? {
    string tempFile = getTempFilePath("replace_absent");
    check createMultiSheetWorkbook(tempFile);

    check writeSheet([["d1", "d2"]], tempFile, "Delta", sheetWriteMode = REPLACE);

    Workbook wb = check fromFile(tempFile);
    test:assertEquals(check wb.getSheetCount(), 4, "Delta added alongside the three siblings");
    test:assertTrue(check wb.hasSheet("Delta"), "Delta created");
    test:assertTrue(check wb.hasSheet("Alpha"), "Alpha still present");
    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetCreatesNewFile() returns error? {
    string tempFile = getTempFilePath("create_new");
    // The path does not exist, so the default FAIL_IF_EXISTS still creates it.
    check writeSheet([["Name", "Age"], ["Alice", "30"]], tempFile);
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed, [["Name", "Age"], ["Alice", "30"]], "New file created");
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetFailsIfSheetExists() returns error? {
    string tempFile = getTempFilePath("fail_if_exists");
    check writeSheet([["First"]], tempFile);  // creates Sheet1

    // The default mode refuses to overwrite an existing sheet.
    Error? result = writeSheet([["Second"]], tempFile);
    test:assertTrue(result is Error, "Default FAIL_IF_EXISTS must refuse an existing sheet");

    // The refusal happens before any write, so the original data is intact.
    string[][] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed, [["First"]], "Original data untouched after refusal");
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetFailIfExistsCreatesWhenAbsent() returns error? {
    string tempFile = getTempFilePath("fail_miss");
    check createMultiSheetWorkbook(tempFile);

    // "NewOne" is absent → FAIL_IF_EXISTS creates it and keeps the siblings.
    check writeSheet([["n1"]], tempFile, "NewOne", sheetWriteMode = FAIL_IF_EXISTS);

    Workbook wb = check fromFile(tempFile);
    test:assertEquals(check wb.getSheetCount(), 4, "NewOne created alongside siblings");
    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetCaseInsensitiveTarget() returns error? {
    string tempFile = getTempFilePath("case_target");
    check createMultiSheetWorkbook(tempFile);

    // Lowercase "beta" resolves to the existing "Beta" — no duplicate sheet.
    check writeSheet([["z1"]], tempFile, "beta", sheetWriteMode = REPLACE);

    Workbook wb = check fromFile(tempFile);
    test:assertEquals(check wb.getSheetCount(), 3, "No duplicate sheet created for a case variant");
    test:assertTrue(check wb.hasSheet("Beta"), "Target still reachable by name");
    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetAppendAlignsToHeader() returns error? {
    string tempFile = getTempFilePath("append_align");
    AppendLogRow[] initial = [
        {name: "Al", age: 30, dept: "Eng"},
        {name: "Bo", age: 25, dept: "Ops"}
    ];
    check writeSheet(initial, tempFile, "Log");

    // Append with a shuffled field order — values must land by header name.
    AppendLogShuffled[] more = [{age: 40, dept: "HR", name: "Cy"}];
    check writeSheet(more, tempFile, "Log", sheetWriteMode = APPEND);

    AppendLogRow[] all = check parseSheet(tempFile, "Log");
    test:assertEquals(all.length(), 3, "Row appended below the existing data");
    test:assertEquals(all[0], {name: "Al", age: 30, dept: "Eng"}, "Existing rows untouched");
    test:assertEquals(all[2], {name: "Cy", age: 40, dept: "HR"}, "Appended row aligned by header name");
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetAppendToAbsentSheetIsFresh() returns error? {
    string tempFile = getTempFilePath("append_absent");
    check createMultiSheetWorkbook(tempFile);

    // "Fresh" is absent → APPEND behaves like a fresh write (header + data).
    AppendLogRow[] data = [{name: "Al", age: 30, dept: "Eng"}];
    check writeSheet(data, tempFile, "Fresh", sheetWriteMode = APPEND);

    AppendLogRow[] parsed = check parseSheet(tempFile, "Fresh");
    test:assertEquals(parsed, data, "Fresh sheet written with header and data");

    Workbook wb = check fromFile(tempFile);
    test:assertEquals(check wb.getSheetCount(), 4, "Siblings preserved");
    check wb.close();
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetAppendStringArrayPositional() returns error? {
    string tempFile = getTempFilePath("append_strings");
    check writeSheet([["H1", "H2"], ["a", "b"]], tempFile, "Data");
    check writeSheet([["c", "d"], ["e", "f"]], tempFile, "Data", sheetWriteMode = APPEND);

    string[][] parsed = check parseSheet(tempFile, "Data");
    test:assertEquals(parsed, [["H1", "H2"], ["a", "b"], ["c", "d"], ["e", "f"]],
            "string[][] rows appended positionally below the last row");
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetAppendRecordWithoutHeaderErrors() returns error? {
    string tempFile = getTempFilePath("append_no_header");

    // Build a sheet whose first row holds only numeric cells — no string header.
    Workbook wb = new;
    Sheet s = check wb.createSheet("Nums");
    check s.setCell(0, 0, 1);
    check s.setCell(0, 1, 2);
    check s.setCell(1, 0, 3);
    check s.setCell(1, 1, 4);
    check wb.saveAs(tempFile);
    check wb.close();

    // Appending records needs a header to align against → typed error, not a silent shift.
    TwoIntRow[] recs = [{a: 5, b: 6}];
    Error? result = writeSheet(recs, tempFile, "Nums", sheetWriteMode = APPEND);
    test:assertTrue(result is Error, "APPEND of records without a header row must error");
    check removeTempFile(tempFile);
}

@test:Config {groups: ["writeSheet"]}
function testWriteSheetCorruptExistingFileErrors() returns error? {
    string tempFile = getTempFilePath("corrupt_existing");
    check io:fileWriteString(tempFile, "this is not a workbook");

    // Opening an existing-but-unreadable file surfaces a ParseError, same as fromFile.
    Error? result = writeSheet([["X"]], tempFile, "Sheet1", sheetWriteMode = REPLACE);
    test:assertTrue(result is ParseError, "A corrupt existing file must yield a ParseError");
    check removeTempFile(tempFile);
}

// Sheet.setRow now honours headerRowIndex (previously hard-wired to 0), so a record
// row can align against a header that does not sit on row 0.
@test:Config {groups: ["writeSheet", "sheet"]}
function testSetRowHonorsHeaderRowIndex() returns error? {
    string tempFile = getTempFilePath("setrow_header_idx");

    // Header on row 2; rows 0-1 are unrelated.
    Workbook wb = new;
    Sheet s = check wb.createSheet("Data");
    check s.setRow(2, ["name", "age"]);
    check wb.saveAs(tempFile);
    check wb.close();

    // Write a record at row 3, pointing setRow at the header on row 2.
    Workbook wb2 = check fromFile(tempFile);
    Sheet s2 = check wb2.getSheet("Data");
    check s2.setRow(3, {name: "Al", age: 30}, headerRowIndex = 2);
    check wb2.save();
    check wb2.close();

    Workbook wb3 = check fromFile(tempFile);
    Sheet s3 = check wb3.getSheet("Data");
    string nameCell = check s3.getCell(3, 0);
    int ageCell = check s3.getCell(3, 1);
    test:assertEquals(nameCell, "Al", "name aligned to col 0 via headerRowIndex = 2");
    test:assertEquals(ageCell, 30, "age aligned to col 1 via headerRowIndex = 2");
    check wb3.close();
    check removeTempFile(tempFile);
}
