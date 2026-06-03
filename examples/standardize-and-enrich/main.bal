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

// Standardize and enrich Excel data.
//
// Reads an Excel file in some source layout, maps it onto a standard schema
// every downstream system understands, enriches each row (region lookup,
// computed total, customer tier, processing date), and writes the standardised
// result as a new Excel file.
//
// The input is read as a byte array and parsed with `xlsx:fromBytes`. This
// example reads the bytes from a local file, creating a sample one first so it
// is self-contained.
//
// Usage:
//   cd examples/standardize-and-enrich
//   bal run
//
// Output: resources/standardized_sales.xlsx

import ballerina/io;
import ballerina/time;
import ballerina/xlsx;

// The source Excel layout (short, source-specific column headers).
type SourceRow record {|
    @xlsx:Name {value: "Cust"}
    string customer;
    @xlsx:Name {value: "Country"}
    string country;
    @xlsx:Name {value: "Item"}
    string item;
    @xlsx:Name {value: "Qty"}
    int quantity;
    @xlsx:Name {value: "Price"}
    decimal price;
|};

// The standard schema every downstream system consumes. The trailing fields
// (region, total, tier, processedOn) are enrichment, not present in the source.
type StandardRow record {|
    @xlsx:Name {value: "Customer"}
    string customer;
    @xlsx:Name {value: "Region"}
    string region;
    @xlsx:Name {value: "Product"}
    string product;
    @xlsx:Name {value: "Quantity"}
    int quantity;
    @xlsx:Name {value: "Unit Price"}
    decimal unitPrice;
    @xlsx:Name {value: "Total"}
    decimal total;
    @xlsx:Name {value: "Tier"}
    string tier;
    @xlsx:Name {value: "Processed On"}
    time:Date processedOn;
|};

public function main() returns error? {
    // (Setup) Create a sample source file so the example is self-contained.
    check createSampleSourceFile("resources/source_sales.xlsx");

    // 1. Read the input file as bytes and parse it with `fromBytes`.
    byte[] inputBytes = check io:fileReadBytes("resources/source_sales.xlsx");
    xlsx:Workbook wb = check xlsx:fromBytes(inputBytes);
    xlsx:Sheet sheet = check wb.getSheet(0);
    SourceRow[] sourceRows = check sheet.getRows();
    check wb.close();
    io:println(string `Read ${sourceRows.length()} source rows.`);

    // 2. Map onto the standard schema and enrich each row.
    time:Date processedOn = today();
    StandardRow[] standardRows = from SourceRow s in sourceRows
        let decimal total = s.price * <decimal>s.quantity
        select {
            customer: s.customer,
            region: regionFor(s.country),
            product: s.item,
            quantity: s.quantity,
            unitPrice: s.price,
            total: total,
            tier: tierFor(total),
            processedOn: processedOn
        };

    // 3. Write the standardised, enriched data to a new Excel file. `writeSheet`
    //    (the one-shot functional API) is the most direct way to land a single
    //    sheet of records on disk.
    check xlsx:writeSheet(standardRows, "resources/standardized_sales.xlsx");
    io:println(string `Wrote ${standardRows.length()} standardised rows to ` +
        string `resources/standardized_sales.xlsx.`);
}

// Map a country code to its sales region (enrichment).
function regionFor(string country) returns string {
    match country {
        "US"|"CA"|"MX" => {
            return "North America";
        }
        "GB"|"DE"|"FR"|"ES"|"IT" => {
            return "Europe";
        }
        "IN"|"CN"|"JP"|"SG" => {
            return "Asia";
        }
        _ => {
            return "Other";
        }
    }
}

// Derive a customer tier from the order total (enrichment).
function tierFor(decimal total) returns string {
    if total >= 1000d {
        return "Gold";
    }
    if total >= 200d {
        return "Silver";
    }
    return "Bronze";
}

// Today's date (UTC) for the processing stamp.
function today() returns time:Date {
    time:Civil now = time:utcToCivil(time:utcNow());
    return {year: now.year, month: now.month, day: now.day};
}

// Create a sample source file to process, so the example is self-contained.
function createSampleSourceFile(string path) returns error? {
    SourceRow[] sample = [
        {customer: "Acme Corp", country: "US", item: "Widget", quantity: 120, price: 4.50},
        {customer: "Globex", country: "DE", item: "Gadget", quantity: 8, price: 29.99},
        {customer: "Initech", country: "IN", item: "Gizmo", quantity: 300, price: 2.10},
        {customer: "Hooli", country: "GB", item: "Sprocket", quantity: 15, price: 88.00}
    ];
    check xlsx:writeSheet(sample, path);
}
