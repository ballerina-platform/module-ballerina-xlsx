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

import ballerina/time;
import ballerina/test;

// =============================================================================
// TEST TYPES
// =============================================================================

type EventWithDate record {|
    string name;
    time:Date eventDate;
|};

type EventWithTime record {|
    string name;
    time:TimeOfDay startTime;
|};

type EventWithDateTime record {|
    string name;
    time:Civil createdAt;
|};

type EventWithNilableDate record {|
    string name;
    time:Date? eventDate;
|};

// =============================================================================
// DATE FIELD TESTS
// =============================================================================

@test:Config {
    groups: ["datetime"]
}
function testWriteAndParseDateField() returns error? {
    // Create records with time:Date fields
    EventWithDate[] events = [
        {name: "Meeting", eventDate: {year: 2026, month: 1, day: 15}},
        {name: "Conference", eventDate: {year: 2026, month: 6, day: 20}}
    ];

    string tempFile = getTempFilePath("datetime_date");
    check writeSheet(events, tempFile);

    // Parse back
    EventWithDate[] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 events");
    test:assertEquals(parsed[0].name, "Meeting", "First event name");
    test:assertEquals(parsed[0].eventDate.year, 2026, "First event year");
    test:assertEquals(parsed[0].eventDate.month, 1, "First event month");
    test:assertEquals(parsed[0].eventDate.day, 15, "First event day");

    test:assertEquals(parsed[1].name, "Conference", "Second event name");
    test:assertEquals(parsed[1].eventDate.year, 2026, "Second event year");
    test:assertEquals(parsed[1].eventDate.month, 6, "Second event month");
    test:assertEquals(parsed[1].eventDate.day, 20, "Second event day");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["datetime"]
}
function testWriteAndParseTimeField() returns error? {
    // Create records with time:TimeOfDay fields
    EventWithTime[] events = [
        {name: "Morning Standup", startTime: {hour: 9, minute: 0, second: 0d}},
        {name: "Team Lunch", startTime: {hour: 12, minute: 30, second: 0d}}
    ];

    string tempFile = getTempFilePath("datetime_time");
    check writeSheet(events, tempFile);

    // Parse back
    EventWithTime[] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 events");
    test:assertEquals(parsed[0].name, "Morning Standup", "First event name");
    test:assertEquals(parsed[0].startTime.hour, 9, "First event hour");
    test:assertEquals(parsed[0].startTime.minute, 0, "First event minute");

    test:assertEquals(parsed[1].name, "Team Lunch", "Second event name");
    test:assertEquals(parsed[1].startTime.hour, 12, "Second event hour");
    test:assertEquals(parsed[1].startTime.minute, 30, "Second event minute");

    check removeTempFile(tempFile);
}

@test:Config {
    groups: ["datetime"]
}
function testWriteAndParseDateTimeField() returns error? {
    // Create records with time:Civil fields
    EventWithDateTime[] events = [
        {
            name: "Webinar",
            createdAt: {year: 2026, month: 3, day: 10, hour: 14, minute: 30, second: 0d}
        },
        {
            name: "Release",
            createdAt: {year: 2026, month: 12, day: 31, hour: 23, minute: 59, second: 59d}
        }
    ];

    string tempFile = getTempFilePath("datetime_civil");
    check writeSheet(events, tempFile);

    // Parse back
    EventWithDateTime[] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 events");

    // First event
    test:assertEquals(parsed[0].name, "Webinar", "First event name");
    test:assertEquals(parsed[0].createdAt.year, 2026, "First event year");
    test:assertEquals(parsed[0].createdAt.month, 3, "First event month");
    test:assertEquals(parsed[0].createdAt.day, 10, "First event day");
    test:assertEquals(parsed[0].createdAt.hour, 14, "First event hour");
    test:assertEquals(parsed[0].createdAt.minute, 30, "First event minute");

    // Second event
    test:assertEquals(parsed[1].name, "Release", "Second event name");
    test:assertEquals(parsed[1].createdAt.year, 2026, "Second event year");
    test:assertEquals(parsed[1].createdAt.month, 12, "Second event month");
    test:assertEquals(parsed[1].createdAt.day, 31, "Second event day");
    test:assertEquals(parsed[1].createdAt.hour, 23, "Second event hour");
    test:assertEquals(parsed[1].createdAt.minute, 59, "Second event minute");

    check removeTempFile(tempFile);
}

// =============================================================================
// NULLABLE DATE TESTS
// =============================================================================

@test:Config {
    groups: ["datetime"]
}
function testWriteAndParseNilableDateField() returns error? {
    // Create records with nilable time:Date fields
    EventWithNilableDate[] events = [
        {name: "Scheduled", eventDate: {year: 2026, month: 5, day: 1}},
        {name: "TBD", eventDate: null}
    ];

    string tempFile = getTempFilePath("datetime_nilable");
    check writeSheet(events, tempFile);

    // Parse back
    EventWithNilableDate[] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 events");
    test:assertEquals(parsed[0].name, "Scheduled", "First event name");
    time:Date? firstDate = parsed[0].eventDate;
    test:assertTrue(firstDate is time:Date, "First event should have date");
    if firstDate is time:Date {
        test:assertEquals(firstDate.year, 2026, "First event year");
        test:assertEquals(firstDate.month, 5, "First event month");
        test:assertEquals(firstDate.day, 1, "First event day");
    }

    test:assertEquals(parsed[1].name, "TBD", "Second event name");
    // Second event's date is null

    check removeTempFile(tempFile);
}

// =============================================================================
// WORKBOOK API TESTS
// =============================================================================

@test:Config {
    groups: ["datetime", "workbook"]
}
function testWorkbookWriteAndParseDateField() returns error? {
    // Test date handling via Workbook/Sheet API
    Workbook wb = new;
    Sheet sheet = check wb.createSheet("Events");

    EventWithDate[] events = [
        {name: "Launch", eventDate: {year: 2026, month: 7, day: 4}}
    ];

    check sheet.putRows(events);

    string tempFile = getTempFilePath("workbook_datetime");
    check wb.saveAs(tempFile);
    check wb.close();

    // Read back via Workbook API
    Workbook wb2 = check fromFile(tempFile);
    Sheet sheet2 = check wb2.getSheet(0);

    EventWithDate[] parsed = check sheet2.getRows();

    test:assertEquals(parsed.length(), 1, "Should have 1 event");
    test:assertEquals(parsed[0].name, "Launch", "Event name");
    test:assertEquals(parsed[0].eventDate.year, 2026, "Event year");
    test:assertEquals(parsed[0].eventDate.month, 7, "Event month");
    test:assertEquals(parsed[0].eventDate.day, 4, "Event day");

    check wb2.close();
    check removeTempFile(tempFile);
}

// =============================================================================
// STRING FALLBACK TESTS
// =============================================================================

@test:Config {
    groups: ["datetime"]
}
function testDateFallbackToStringWhenTargetIsString() returns error? {
    // When target type is string, date cells should fall back to ISO format
    EventWithDate[] events = [
        {name: "Test", eventDate: {year: 2026, month: 1, day: 1}}
    ];

    string tempFile = getTempFilePath("datetime_string_fallback");
    check writeSheet(events, tempFile);

    // Parse as string[][] (should get ISO date string)
    string[][] parsed = check parseSheet(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have header + 1 data row");
    test:assertEquals(parsed[0][0], "name", "First header");
    test:assertEquals(parsed[0][1], "eventDate", "Second header");
    test:assertEquals(parsed[1][0], "Test", "Data name");
    // The date should be in ISO format or similar string representation
    test:assertTrue(parsed[1][1].length() > 0, "Date should have string value");

    check removeTempFile(tempFile);
}

// =============================================================================
// Sub-second precision round-trip tests
// =============================================================================
// `time:Civil.second` and `time:TimeOfDay.second` are `decimal` in the Ballerina
// time module. The native layer must preserve nano precision on both read and
// write paths via BigDecimal arithmetic — integer-second math would silently
// truncate fractional seconds.

type EventWithCivilSubSecond record {|
    string name;
    time:Civil ts;
|};

@test:Config {groups: ["datetime"]}
function testCivilSubSecondPrecision() returns error? {
    // Civil values round-trip at microsecond precision, not nanosecond. Excel
    // stores all datetime cells as `double` serials (~17 significant decimal
    // digits); when the date integer is included (e.g., year 2026 → ~46,000),
    // only ~12 digits remain for the fractional day — ~microsecond resolution.
    // This is a fundamental limit of the Excel storage format, not our module.
    // Pure time-of-day values (no date) preserve full nanosecond precision —
    // see testTimeOfDaySubSecondPrecision.
    time:Civil exact = {
        year: 2026, month: 5, day: 27,
        hour: 12, minute: 30, second: 1.123457d,
        utcOffset: {hours: 0, minutes: 0}
    };
    EventWithCivilSubSecond[] events = [{name: "Tick", ts: exact}];
    string tempFile = getTempFilePath("civil_subsec");
    check writeSheet(events, tempFile);

    EventWithCivilSubSecond[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 1);
    test:assertEquals(parsed[0].ts.year, 2026);
    test:assertEquals(parsed[0].ts.month, 5);
    test:assertEquals(parsed[0].ts.day, 27);
    test:assertEquals(parsed[0].ts.hour, 12);
    test:assertEquals(parsed[0].ts.minute, 30);
    test:assertEquals(parsed[0].ts.second, 1.123457d);

    check removeTempFile(tempFile);
}

@test:Config {groups: ["datetime"]}
function testTimeOfDaySubSecondPrecision() returns error? {
    time:TimeOfDay exact = {hour: 14, minute: 45, second: 30.987654321d};
    EventWithTime[] events = [{name: "Beat", startTime: exact}];
    string tempFile = getTempFilePath("timeofday_subsec");
    check writeSheet(events, tempFile);

    EventWithTime[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 1);
    test:assertEquals(parsed[0].startTime.hour, 14);
    test:assertEquals(parsed[0].startTime.minute, 45);
    test:assertEquals(parsed[0].startTime.second, 30.987654321d);

    check removeTempFile(tempFile);
}

// =============================================================================
// 1904 epoch parsing
// =============================================================================
// Workbooks produced by Excel for Mac legacy / `openpyxl(epoch1904=True)` use a
// 1904 reference date instead of the default 1900 one. The native parser must
// honour `Workbook.isDate1904()` rather than always assuming 1900.

type Date1904Row record {|
    time:Date col0;
|};

@test:Config {groups: ["datetime"]}
function testParseDate1904Workbook() returns error? {
    // dates_1904.xlsx is generated programmatically in @test:BeforeSuite with the
    // 1904 epoch flag set. Cell (0, 0) holds the numeric serial 44708.
    //
    // Under the 1904 epoch, 1904-01-01 + 44708 days = 2026-05-28.
    // Under the 1900 epoch, the same serial would land in mid-2022 — wildly off.
    // The assertion below is correct ONLY if Workbook.isDate1904() is honoured.
    Workbook wb = check fromFile(TEST_DATA_DIR + "dates_1904.xlsx");
    Sheet sheet = check wb.getSheet(0);

    // The cell isn't date-formatted in this fixture (a raw int was written), so
    // we exercise the 1904 branch in CellConverter.convertByType via the
    // target-type-triggered path: parsing into a record whose field is typed
    // time:Date drives convertSerialToLocalDate with isDate1904 = true.
    // headerRowIndex: () puts the reader in headerless mode so column 0 is
    // exposed as "col0" — matching the record's field name.
    Date1904Row row = check sheet.getRow(0, {headerRowIndex: ()});
    test:assertEquals(row.col0, {year: 2026, month: 5, day: 28},
            "1904-epoch serial 44708 should map to 2026-05-28");

    check wb.close();
}

// =============================================================================
// Timezone-independent round-trip
// =============================================================================
// All naive cell conversions go through ZoneOffset.UTC; the same workbook
// produces the same time:Date / time:Civil values regardless of the system
// timezone of the machine running the tests.

@test:Config {groups: ["datetime"]}
function testTimezoneIndependentDateRoundTrip() returns error? {
    EventWithDate[] events = [
        {name: "Sunrise", eventDate: {year: 2026, month: 5, day: 27}},
        {name: "Sunset", eventDate: {year: 2026, month: 12, day: 31}}
    ];
    string tempFile = getTempFilePath("tz_indep_date");
    check writeSheet(events, tempFile);

    EventWithDate[] parsed = check parseSheet(tempFile);
    test:assertEquals(parsed.length(), 2);
    test:assertEquals(parsed[0].eventDate, {year: 2026, month: 5, day: 27});
    test:assertEquals(parsed[1].eventDate, {year: 2026, month: 12, day: 31});

    check removeTempFile(tempFile);
}

// =============================================================================
// Datetime cells preserve their time component when read as string
// =============================================================================
// convertToString used to drop the time portion of any date-formatted cell —
// every datetime collapsed to "yyyy-MM-dd". The fix branches on the serial's
// integer + fractional parts so datetime cells keep "yyyy-MM-dd HH:mm:ss" and
// time-only cells become "HH:mm:ss", while still bypassing POI's
// timezone-dependent DataFormatter to preserve cross-machine determinism.

type CivilEvent record {|
    string name;
    time:Civil ts;
|};

@test:Config {groups: ["datetime"]}
function testDatetimeCellReadAsStringPreservesTime() returns error? {
    string tempFile = getTempFilePath("datetime_as_string");
    CivilEvent[] data = [{
        name: "Event",
        ts: {
            year: 2026, month: 5, day: 28,
            hour: 14, minute: 30, second: 45d,
            utcOffset: {hours: 0, minutes: 0}
        }
    }];
    check writeSheet(data, tempFile);

    string[][] rows = check parseSheet(tempFile);
    test:assertEquals(rows.length(), 2, "Expected header row + 1 data row");
    string tsCell = rows[1][1];
    test:assertEquals(tsCell, "2026-05-28 14:30:45",
            "Datetime cell must preserve the full date and time as 'yyyy-MM-dd HH:mm:ss'");

    check removeTempFile(tempFile);
}
