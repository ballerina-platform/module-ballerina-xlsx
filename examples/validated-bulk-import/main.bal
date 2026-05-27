// Validated Bulk Import
//
// Pattern: a downstream pipeline accepts a partner's XLSX upload. Real-world
// data is messy — some rows fail business rules. This example shows how to
// keep the pipeline clean while preserving full visibility into rejected rows:
//
//   1. Apply Ballerina `@constraint` annotations to the target record type.
//   2. Parse with `enableConstraintValidation` + `FailSafeOptions` writing
//      rejected rows to a JSON log file. Valid rows flow through; bad rows
//      land in the log with their raw cell values and reason for rejection.
//   3. `caseInsensitiveHeaders` makes the import tolerant of partner-side
//      capitalisation choices.

import ballerina/constraint;
import ballerina/file;
import ballerina/io;
import ballerinax/xlsx;

type Customer record {|
    @xlsx:Name {value: "Full Name"}
    @constraint:String {minLength: 1, maxLength: 64}
    string fullName;

    @xlsx:Name {value: "Email"}
    @constraint:String {pattern: re `^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$`}
    string email;

    @xlsx:Name {value: "Age"}
    @constraint:Int {minValue: 18, maxValue: 120}
    int age;

    @xlsx:Name {value: "Country"}
    string country;
|};

public function main() returns error? {
    string inputPath = "resources/customers.xlsx";
    string errorLog = "resources/import_errors.json";
    if !check file:test("resources", file:EXISTS) {
        check file:createDir("resources");
    }

    // Simulate the partner-supplied file. Rows 3, 4, 5, 7 deliberately violate
    // the constraints declared on Customer above. Note the deliberately-mixed
    // header capitalisation ("FULL NAME" / "email") — `caseInsensitiveHeaders`
    // below makes the import tolerant of it.
    string[][] partnerData = [
        ["FULL NAME", "email", "Age", "country"],
        ["Alice Adams", "alice@example.com", "32", "US"],
        ["Bob Banner", "bob@example.com", "47", "UK"],
        ["Carol Chen", "not-an-email", "29", "SG"],
        ["", "diana@example.com", "41", "FR"],
        ["Eve Edwards", "eve@example.com", "12", "JP"],
        ["Frank Foster", "frank@example.com", "55", "DE"],
        ["Grace Greene", "grace@example.com", "190", "CA"],
        ["Henry Holmes", "henry@example.com", "38", "AU"],
        ["Ivy Ito", "ivy@example.com", "26", "JP"],
        ["Jack Jones", "jack@example.com", "34", "US"]
    ];
    check xlsx:writeSheet(partnerData, inputPath);
    io:println(string `Partner supplied ${partnerData.length() - 1} customer rows (some invalid).`);

    // Parse with constraint validation + file-based error logging. Bad rows
    // are skipped and recorded; the caller receives only the clean records.
    xlsx:ParseOptions opts = {
        enableConstraintValidation: true,
        caseInsensitiveHeaders: true,
        failSafe: {
            enableConsoleLogs: false,
            fileOutputMode: {
                filePath: errorLog,
                contentType: xlsx:RAW_AND_METADATA,
                fileWriteOption: xlsx:OVERWRITE
            }
        }
    };
    Customer[] accepted = check xlsx:parseSheet(inputPath, 0, opts);

    io:println("");
    io:println(string `Accepted ${accepted.length()} clean rows:`);
    foreach Customer c in accepted {
        io:println(string `  - ${c.fullName} <${c.email}> (age ${c.age}, ${c.country})`);
    }

    // The rejected rows are now in the error log with their original cell
    // values plus the validation reason — exactly what the data ops team
    // needs to follow up with the partner.
    io:println("");
    io:println(string `Rejected rows logged to ${errorLog}:`);
    string logContent = check io:fileReadString(errorLog);
    io:println(logContent);
}
