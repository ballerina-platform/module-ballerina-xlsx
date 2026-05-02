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
    check write(events, tempFile);

    // Parse back
    EventWithDate[] parsed = check parse(tempFile);

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
    check write(events, tempFile);

    // Parse back
    EventWithTime[] parsed = check parse(tempFile);

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
    check write(events, tempFile);

    // Parse back
    EventWithDateTime[] parsed = check parse(tempFile);

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
    check write(events, tempFile);

    // Parse back
    EventWithNilableDate[] parsed = check parse(tempFile);

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
    Workbook wb = check createWorkbook();
    Sheet sheet = check wb.createSheet("Events");

    EventWithDate[] events = [
        {name: "Launch", eventDate: {year: 2026, month: 7, day: 4}}
    ];

    check sheet.putRows(events);

    string tempFile = getTempFilePath("workbook_datetime");
    check wb.saveAs(tempFile);
    check wb.close();

    // Read back via Workbook API
    Workbook wb2 = check openFile(tempFile);
    Sheet sheet2 = check wb2.getSheetByIndex(0);

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
    check write(events, tempFile);

    // Parse as string[][] (should get ISO date string)
    string[][] parsed = check parse(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have header + 1 data row");
    test:assertEquals(parsed[0][0], "name", "First header");
    test:assertEquals(parsed[0][1], "eventDate", "Second header");
    test:assertEquals(parsed[1][0], "Test", "Data name");
    // The date should be in ISO format or similar string representation
    test:assertTrue(parsed[1][1].length() > 0, "Date should have string value");

    check removeTempFile(tempFile);
}
