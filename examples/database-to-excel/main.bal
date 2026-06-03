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

// Database -> Excel export.
//
// Reads rows from a database, maps them onto the column layout a downstream
// consumer expects, builds an Excel workbook with the Workbook API, serialises
// it to a byte array with `Workbook.toBytes()`, and writes those bytes to a file.
//
// The database is an in-memory H2 instance seeded on the fly, so the example
// runs with no external database; point `jdbc:Client` at your own database to
// use real data.
//
// Usage:
//   cd examples/database-to-excel
//   bal run
//
// Output: resources/orders.xlsx

import ballerina/io;
import ballerina/sql;
import ballerinax/java.jdbc;
import ballerina/xlsx;

// The shape of a row in the database.
type OrderRecord record {|
    int id;
    string customer;
    string sku;
    int quantity;
    decimal unitPrice;
|};

// The consumer's expected Excel layout. `@xlsx:Name` is the mapping: each field
// is written under the exact column header the consumer reads.
type VendorOrder record {|
    @xlsx:Name {value: "Order ID"}
    int id;
    @xlsx:Name {value: "Customer Name"}
    string customer;
    @xlsx:Name {value: "Product SKU"}
    string sku;
    @xlsx:Name {value: "Quantity"}
    int quantity;
    @xlsx:Name {value: "Unit Price (USD)"}
    decimal unitPrice;
    @xlsx:Name {value: "Line Total (USD)"}
    decimal lineTotal;
|};

public function main() returns error? {
    // 1. Read the source rows from the database.
    OrderRecord[] orders = check loadOrdersFromDatabase();
    io:println(string `Read ${orders.length()} orders from the database.`);

    // 2. Map database rows onto the consumer's schema, deriving the line total.
    VendorOrder[] vendorRows = from OrderRecord o in orders
        select {
            id: o.id,
            customer: o.customer,
            sku: o.sku,
            quantity: o.quantity,
            unitPrice: o.unitPrice,
            lineTotal: o.unitPrice * <decimal>o.quantity
        };

    // 3. Build the workbook and serialise it to a byte array.
    xlsx:Workbook wb = new;
    xlsx:Sheet sheet = check wb.createSheet("Orders");
    check sheet.putRows(vendorRows);
    byte[] xlsxBytes = check wb.toBytes();
    check wb.close();

    // 4. Write the bytes to a file.
    check io:fileWriteBytes("resources/orders.xlsx", xlsxBytes);
    io:println(string `Wrote ${vendorRows.length()} rows to ` +
        string `resources/orders.xlsx (${xlsxBytes.length()} bytes).`);
}

// Seed an in-memory H2 database with sample orders and read them back. The
// `jdbc:h2:mem:` URL keeps everything in process, so the example needs no
// external database — only Ballerina and the xlsx module.
function loadOrdersFromDatabase() returns OrderRecord[]|error {
    jdbc:Client db = check new ("jdbc:h2:mem:vendordb;DB_CLOSE_DELAY=-1");

    _ = check db->execute(`CREATE TABLE orders (
        id INT PRIMARY KEY,
        customer VARCHAR(100),
        sku VARCHAR(50),
        quantity INT,
        unit_price DECIMAL(10, 2)
    )`);

    OrderRecord[] seed = [
        {id: 1001, customer: "Acme Corp", sku: "SKU-100", quantity: 10, unitPrice: 19.99},
        {id: 1002, customer: "Globex", sku: "SKU-205", quantity: 4, unitPrice: 149.50},
        {id: 1003, customer: "Initech", sku: "SKU-310", quantity: 25, unitPrice: 8.75},
        {id: 1004, customer: "Umbrella", sku: "SKU-410", quantity: 2, unitPrice: 1200.00}
    ];
    foreach OrderRecord o in seed {
        _ = check db->execute(`INSERT INTO orders (id, customer, sku, quantity, unit_price)
            VALUES (${o.id}, ${o.customer}, ${o.sku}, ${o.quantity}, ${o.unitPrice})`);
    }

    // `unit_price` is aliased to match the camelCase record field (column-name
    // matching is case-insensitive but does not bridge snake_case). The row type
    // is passed explicitly because it cannot be inferred from the query clause.
    stream<OrderRecord, sql:Error?> rs = db->query(
        `SELECT id, customer, sku, quantity, unit_price AS unitPrice FROM orders ORDER BY id`,
        OrderRecord);
    OrderRecord[] orders = check from OrderRecord o in rs
        select o;

    check db.close();
    return orders;
}
