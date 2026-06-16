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

// =============================================================================
// SHARED TEST RECORD TYPES
// =============================================================================
// These record types are used across multiple test files for consistent testing.

// Basic employee record for simple tests
type Employee record {|
    string name;
    int age;
    string department;
|};

// Record with various data types for type conversion testing
type TypeVariety record {|
    string text;
    int number;
    decimal amount;
    boolean flag;
|};

// Record with nilable fields for optional value testing
type NilableRecord record {|
    string name;
    int? age;
    string? department;
|};

// Record with @xlsx:Name annotations for header mapping tests
type AnnotatedEmployee record {|
    @Name {value: "First Name"}
    string firstName;
    @Name {value: "Employee ID"}
    int id;
    @Name {value: "Department Name"}
    string department;
|};

// Record for testing numeric type conversions
type NumericTypes record {|
    int intValue;
    decimal decimalValue;
|};

// Simple record for basic write/read tests
type Person record {|
    string name;
    int age;
    boolean active;
|};

// Record for testing wide data (many columns)
type WideRecord record {|
    string col1;
    string col2;
    string col3;
    string col4;
    string col5;
    string col6;
    string col7;
    string col8;
    string col9;
    string col10;
|};

// =============================================================================
// EXPECTED DATA CONSTANTS
// =============================================================================
// Known expected values for generated test files. These must match the data
// generated in test_utils.bal setupTestData().

// Expected data for simple.xlsx
final string[][] EXPECTED_SIMPLE_DATA = [
    ["Name", "Age", "City"],
    ["John", "30", "New York"],
    ["Jane", "25", "Los Angeles"],
    ["Bob", "35", "Chicago"]
];

// Expected employees for employees.xlsx
final Employee[] EXPECTED_EMPLOYEES = [
    {name: "John Doe", age: 30, department: "Engineering"},
    {name: "Jane Smith", age: 28, department: "Marketing"},
    {name: "Bob Johnson", age: 35, department: "Sales"}
];

// Expected data for multi_sheet.xlsx - Sheet1
final string[][] EXPECTED_SHEET1_DATA = [
    ["A1", "B1"],
    ["A2", "B2"]
];

// Expected data for multi_sheet.xlsx - Sheet2
final string[][] EXPECTED_SHEET2_DATA = [
    ["X1", "Y1"],
    ["X2", "Y2"]
];

// Expected data for multi_sheet.xlsx - Sheet3
final string[][] EXPECTED_SHEET3_DATA = [
    ["P1", "Q1"],
    ["P2", "Q2"]
];

// Expected formula values (cached results)
final string[][] EXPECTED_FORMULA_CACHED = [
    ["A", "B", "Sum"],
    ["10", "20", "30"],   // =A2+B2 = 10+20 = 30
    ["15", "25", "40"]    // =A3+B3 = 15+25 = 40
];

// Expected formula text (formula strings)
final string[][] EXPECTED_FORMULA_TEXT = [
    ["A", "B", "Sum"],
    ["10", "20", "=A2+B2"],
    ["15", "25", "=A3+B3"]
];

// =============================================================================
// DATA PROJECTION TEST TYPES
// =============================================================================

// Record with optional fields for nilAsOptionalField testing
// Uses @xlsx:Name to match simple.xlsx headers (Name, Age, City)
type RecordWithOptionalFields record {|
    @Name {value: "Name"}
    string name;
    @Name {value: "Age"}
    int? age;        // nilable
    @Name {value: "City"}
    string? city;    // nilable
|};

// Record with required field that won't be in the sheet (for absentAsNilableType testing)
// Uses @xlsx:Name to match simple.xlsx headers (Name, Age, City)
// extraField has no matching column in simple.xlsx
type RecordWithExtraField record {|
    @Name {value: "Name"}
    string name;
    @Name {value: "Age"}
    int age;
    string? extraField;  // nilable, won't have matching column
|};

// Record for strict mode testing (all required fields)
// department has no matching column in simple.xlsx
type StrictModeRecord record {|
    @Name {value: "Name"}
    string name;
    @Name {value: "Age"}
    int age;
    string department;  // required, no matching column in simple.xlsx
|};

// =============================================================================
// CASE-INSENSITIVE HEADER TEST TYPES
// =============================================================================

// Record for case-insensitive header testing
// Field names are lowercase, test data has mixed-case headers (NAME, AGE, Department)
type CaseTestEmployee record {|
    string name;       // Should match "NAME" when case-insensitive
    int age;           // Should match "AGE" when case-insensitive
    string department; // Should match "Department" when case-insensitive
|};

// =============================================================================
// ERROR TYPE PRESERVATION TEST TYPES
// =============================================================================

// Record for testing error type preservation - simple record with required int field
type ErrorTypeTestRecord record {|
    string name;
    int age;  // Required - invalid values should cause TypeConversionError
|};

// =============================================================================
// BLANK CELL HANDLING TEST TYPES
// =============================================================================

// Record with required non-nilable age field
type RequiredAgeRecord record {|
    string name;
    int age;  // Required, non-nilable - blank cell should cause error
|};

// Record with optional age field
type OptionalAgeRecord record {|
    string name;
    int age?;  // Optional - blank cell should succeed
|};

// Record with nilable age field
type NilableAgeRecord record {|
    string name;
    int? age;  // Nilable - blank cell should result in null
|};

// =============================================================================
// ROW WRAPPER TEST TYPES
// =============================================================================

// Simple record for Row wrapper tests
type SimpleData record {|
    string name;
    int value;
|};

// =============================================================================
// HEADER-LESS PARSING TEST TYPES
// =============================================================================

// Record for header-less parsing - uses col0, col1, col2 as field names
type HeaderlessRecord record {|
    string col0;  // First column
    string col1;  // Second column
|};

// =============================================================================
// EXTRA-COLUMN CAPTURE TEST TYPES
// =============================================================================

// Record that only defines some fields, with a CellValue rest descriptor so columns
// beyond the declared fields are captured into the rest field. The rest type must be
// CellValue (not the open-record `anydata` default) to satisfy the `Row` bound.
type OpenEmployee record {|
    string name;
    int age;
    // No 'department' field defined - extra columns populate the rest field
    CellValue...;
|};

// Declares only some of natural_types.xlsx's columns; the remaining typed columns
// (decimalCol, dateCol, datetimeCol) fall to the CellValue rest field, which must
// keep their natural Ballerina types rather than collapsing to strings.
type PartialNaturalRow record {|
    int intCol;
    boolean boolCol;
    CellValue...;
|};

// Maps the genuinely-typed scalar columns of natural_types.xlsx (int, decimal,
// boolean) onto strongly-typed record fields to exercise the typed-cell →
// typed-record binding path.
type NaturalTypedRow record {|
    int intCol;
    decimal decimalCol;
    boolean boolCol;
|};

// =============================================================================
// HELPER CONSTANTS
// =============================================================================

// Test data directory path
const string TEST_DATA_DIR = "tests/resources/testdata/";
