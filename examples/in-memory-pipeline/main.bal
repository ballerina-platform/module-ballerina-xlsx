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

// In-Memory Pipeline
//
// Pattern: a microservice receives an XLSX as a byte array (HTTP upload,
// queue payload, S3 download — any byte source), applies a transformation,
// and returns the modified workbook as a byte array. No disk I/O at any
// point. This is the primary use case the byte-array API surface is for.
//
// The transformation here is a 10% tax surcharge applied to every row of an
// invoice sheet. The handler is a pure `byte[] -> byte[]` function that drops
// straight into any transport (HTTP handler, queue consumer, batch worker)
// without modification.

import ballerina/io;
import ballerina/xlsx;

type Invoice record {|
    string itemCode;
    string description;
    decimal amount;
|};

public function main() returns error? {
    // Build an "incoming" workbook entirely in memory and serialise to bytes.
    // Simulates an XLSX arriving over HTTP / a queue / from cloud storage.
    byte[] incomingBytes = check buildIncomingPayload();
    io:println(string `Received ${incomingBytes.length()} bytes of incoming workbook.`);

    // The actual handler — pure byte[] in, byte[] out, no filesystem touched.
    byte[] outgoingBytes = check applyTaxSurcharge(incomingBytes, 0.10d);
    io:println(string `Returning ${outgoingBytes.length()} bytes of modified workbook.`);
    io:println("");

    // Verify: load the outgoing bytes back into a workbook and read the
    // transformed amounts.
    xlsx:Workbook reopened = check xlsx:fromBytes(outgoingBytes);
    xlsx:Sheet sheet = check reopened.getSheet(0);
    Invoice[] result = check sheet.getRows();
    io:println("Transformed invoices (10% surcharge applied):");
    foreach Invoice inv in result {
        io:println(string `  ${inv.itemCode}  ${inv.description}: $${formatUsd(inv.amount)}`);
    }
    check reopened.close();
}

// Round a positive decimal to cents (2 decimal places) using Ballerina's
// native banker's rounding on the int cast — ties resolve to the even
// neighbor (the standard for monetary rounding, matching IEEE 754 half-even).
// A real money handler rounds during the transformation rather than letting
// raw multiplication leak fractional cents into the spreadsheet.
function roundToCents(decimal amount) returns decimal {
    int cents = <int>(amount * 100d);
    return <decimal>cents / 100d;
}

// Format a decimal as a 2-decimal currency string. Ballerina's decimal
// toString() drops trailing zeros — "137.5" instead of "137.50" — which is
// fine mathematically but reads wrong for money.
function formatUsd(decimal amount) returns string {
    string s = amount.toString();
    int? dotIdx = s.indexOf(".");
    if dotIdx is () {
        return s + ".00";
    }
    int fracLen = s.length() - dotIdx - 1;
    if fracLen == 1 {
        return s + "0";
    }
    return s;
}

// Apply a per-row surcharge to the `amount` column of an in-memory workbook
// without ever touching disk. This is the function shape an HTTP service or
// queue consumer would expose — `byte[] -> byte[]|error`.
function applyTaxSurcharge(byte[] input, decimal rate) returns byte[]|error {
    xlsx:Workbook wb = check xlsx:fromBytes(input);
    xlsx:Sheet sheet = check wb.getSheet(0);

    // Discover the data extent rather than assuming a fixed shape — real
    // partner data rarely fits a hard-coded assumption.
    xlsx:CellRange? used = check sheet.getUsedCellRange();
    if used is () {
        return error("Empty sheet — nothing to transform");
    }

    // Locate the amount column by reading the header row directly.
    int amountCol = -1;
    foreach int c in used.firstColumnIndex ... used.lastColumnIndex {
        anydata header = check sheet.getCell(used.firstRowIndex, c);
        if header.toString() == "amount" {
            amountCol = c;
            break;
        }
    }
    if amountCol < 0 {
        return error("Incoming sheet has no 'amount' column");
    }

    // Walk data rows (everything below the header) and apply the surcharge.
    // setCell uses (row, col) — the natural form inside a loop.
    foreach int r in (used.firstRowIndex + 1) ... used.lastRowIndex {
        anydata current = check sheet.getCell(r, amountCol);
        decimal raw;
        if current is decimal {
            raw = current * (1d + rate);
        } else if current is int {
            raw = <decimal>current * (1d + rate);
        } else {
            continue;
        }
        check sheet.setCell(r, amountCol, roundToCents(raw));
    }

    // Record an audit trail on a separate sheet. Keeping the metadata off the
    // data sheet means downstream parsers don't have to filter marker rows
    // out of the Invoices data. setCellByAddress uses A1 notation — the
    // natural form when writing known fixed cells.
    xlsx:Sheet auditSheet = check wb.createSheet("Audit");
    check auditSheet.setCellByAddress("A1", "Tax surcharge applied");
    check auditSheet.setCellByAddress("B1", string `${rate * 100d}%`);

    return wb.toBytes();
}

// Build the "incoming" workbook from in-memory records. Returns the byte
// payload as if it came in over the wire.
function buildIncomingPayload() returns byte[]|error {
    Invoice[] invoices = [
        {itemCode: "SKU-100", description: "Widget", amount: 19.99},
        {itemCode: "SKU-200", description: "Gadget", amount: 49.50},
        {itemCode: "SKU-300", description: "Gizmo", amount: 8.25},
        {itemCode: "SKU-400", description: "Sprocket", amount: 125.00}
    ];
    xlsx:Workbook wb = new;
    xlsx:Sheet sheet = check wb.createSheet("Invoices");
    check sheet.putRows(invoices);
    byte[] bytes = check wb.toBytes();
    check wb.close();
    return bytes;
}
