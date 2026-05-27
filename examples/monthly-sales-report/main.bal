// Monthly Sales Report
//
// Builds a multi-sheet workbook from in-memory data, embeds an Excel Table on
// the "Orders" sheet, writes a regional summary on a second sheet, then
// reopens the file and queries it back. Demonstrates the Workbook + Sheet +
// Table tier of the xlsx API together with target-type-driven `time:Date`
// binding.

import ballerina/file;
import ballerina/io;
import ballerina/time;
import ballerinax/xlsx;

type Order record {|
    @xlsx:Name {value: "Order ID"}
    string id;
    @xlsx:Name {value: "Order Date"}
    time:Date orderDate;
    @xlsx:Name {value: "Customer"}
    string customer;
    @xlsx:Name {value: "Region"}
    string region;
    @xlsx:Name {value: "Amount (USD)"}
    decimal amount;
|};

type RegionSummary record {|
    @xlsx:Name {value: "Region"}
    string region;
    @xlsx:Name {value: "Order Count"}
    int orderCount;
    @xlsx:Name {value: "Total Revenue (USD)"}
    decimal totalRevenue;
|};

public function main() returns error? {
    string path = "resources/sales_q1.xlsx";
    if !check file:test("resources", file:EXISTS) {
        check file:createDir("resources");
    }

    Order[] orders = [
        {id: "ORD-001", orderDate: {year: 2026, month: 1, day: 12}, customer: "Acme Corp", region: "EMEA", amount: 4250.00},
        {id: "ORD-002", orderDate: {year: 2026, month: 1, day: 18}, customer: "Globex", region: "AMER", amount: 1875.50},
        {id: "ORD-003", orderDate: {year: 2026, month: 1, day: 24}, customer: "Initech", region: "AMER", amount: 9320.00},
        {id: "ORD-004", orderDate: {year: 2026, month: 2, day: 3}, customer: "Umbrella", region: "EMEA", amount: 2100.75},
        {id: "ORD-005", orderDate: {year: 2026, month: 2, day: 9}, customer: "Tyrell", region: "APAC", amount: 7600.00},
        {id: "ORD-006", orderDate: {year: 2026, month: 2, day: 14}, customer: "Cyberdyne", region: "AMER", amount: 5430.20},
        {id: "ORD-007", orderDate: {year: 2026, month: 2, day: 22}, customer: "Hooli", region: "APAC", amount: 3210.00},
        {id: "ORD-008", orderDate: {year: 2026, month: 3, day: 2}, customer: "Pied Piper", region: "AMER", amount: 1140.00},
        {id: "ORD-009", orderDate: {year: 2026, month: 3, day: 11}, customer: "Stark Industries", region: "EMEA", amount: 8870.50},
        {id: "ORD-010", orderDate: {year: 2026, month: 3, day: 18}, customer: "Wayne Enterprises", region: "AMER", amount: 6720.00},
        {id: "ORD-011", orderDate: {year: 2026, month: 3, day: 25}, customer: "Soylent Corp", region: "APAC", amount: 4015.00},
        {id: "ORD-012", orderDate: {year: 2026, month: 3, day: 29}, customer: "Massive Dynamic", region: "EMEA", amount: 5550.00}
    ];

    // Build the workbook from scratch using the Workbook API.
    xlsx:Workbook wb = check new;

    // Orders sheet — wrapped in a real Excel Table via createTableFromData,
    // so the data range is registered as a structured Table on open.
    xlsx:Sheet ordersSheet = check wb.createSheet("Orders");
    _ = check ordersSheet.createTableFromData("OrdersTable", orders);
    io:println(string `Wrote ${orders.length()} orders to the 'Orders' sheet (Excel Table: OrdersTable).`);

    // Aggregate by region. Standard Ballerina query expression — the xlsx
    // module does not impose its own query DSL.
    string[] regions = ["AMER", "EMEA", "APAC"];
    RegionSummary[] summary = [];
    foreach string r in regions {
        int count = 0;
        decimal total = 0d;
        foreach Order o in orders {
            if o.region == r {
                count += 1;
                total += o.amount;
            }
        }
        summary.push({region: r, orderCount: count, totalRevenue: total});
    }

    // Summary sheet — write the aggregated records, then post-process by
    // adding a "% of Total" column using setColumn (data lives in column D).
    xlsx:Sheet summarySheet = check wb.createSheet("Summary");
    check summarySheet.putRows(summary);

    decimal grandTotal = 0d;
    foreach RegionSummary s in summary {
        grandTotal += s.totalRevenue;
    }
    // setCell writes any single cell (including the header row 0). setColumn
    // writes data values starting at row 1 — so we set the header explicitly
    // and feed setColumn the data values only.
    check summarySheet.setCell(0, 3, "% of Total");
    anydata[] pctValues = [];
    foreach RegionSummary s in summary {
        decimal pct = (s.totalRevenue * 100d) / grandTotal;
        // Round to 2 decimal places. Ballerina decimal arithmetic preserves
        // arbitrary precision, so without rounding the cell would hold 30+
        // digits — mathematically correct but ugly in a spreadsheet. The
        // <int>decimal cast uses banker's rounding (half-even).
        decimal rounded = <decimal>(<int>(pct * 100d)) / 100d;
        pctValues.push(string `${rounded}%`);
    }
    check summarySheet.setColumn(3, pctValues);

    check wb.saveAs(path);
    check wb.close();
    io:println(string `Saved workbook to ${path}.`);
    io:println("");

    // Reopen the file and inspect it via the Workbook + Table APIs.
    xlsx:Workbook reopened = check new (path);

    string[] sheetNames = reopened.getSheetNames();
    string joined = "";
    foreach string n in sheetNames {
        joined = joined == "" ? n : joined + ", " + n;
    }
    io:println(string `Reopened workbook has ${sheetNames.length()} sheets: ${joined}`);

    xlsx:Table[] allTables = check reopened.getAllTables();
    io:println(string `Workbook contains ${allTables.length()} table(s).`);

    xlsx:Table ordersTable = check reopened.getTable("OrdersTable");
    string[] headers = ordersTable.getHeaders();
    string headersJoined = "";
    foreach string h in headers {
        headersJoined = headersJoined == "" ? h : headersJoined + " | " + h;
    }
    io:println(string `Table '${ordersTable.getName()}' columns: ${headersJoined}`);

    // Read all orders back through the Table API — headers are excluded.
    Order[] readBack = check ordersTable.getRows();
    Order topOrder = readBack[0];
    foreach Order o in readBack {
        if o.amount > topOrder.amount {
            topOrder = o;
        }
    }
    io:println(string `Top order: ${topOrder.id} - ${topOrder.customer} (${topOrder.region}) for $${topOrder.amount}`);

    // Spot-check a single cell directly without parsing the whole sheet.
    xlsx:Sheet ordersSheetReopened = check reopened.getSheet("Orders");
    anydata firstCustomer = check ordersSheetReopened.getCell(1, 2); // with 0-based indices, this means 1 -> 2, 2 -> C, i.e. C2 cell. 
    io:println(string `Spot check via Sheet.getCell(1, 2): first customer is "${firstCustomer.toString()}".`);

    check reopened.close();
}
