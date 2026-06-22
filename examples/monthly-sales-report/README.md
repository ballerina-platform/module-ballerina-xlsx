# Monthly Sales Report

Builds a multi-sheet workbook from scratch with the Workbook + Sheet + Table APIs, then reopens the file
and queries it back. An "Orders" sheet is wrapped in a real Excel Table (`createTableFromData`); a
"Summary" sheet is aggregated by region with a computed "% of Total" column added via `setCell` (header)
and `setColumn` (values). Order dates use target-type-driven `time:Date` binding.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

1. Creates a workbook in memory (`new`), writes 12 orders to the "Orders" sheet as an Excel Table named
   `OrdersTable`, and writes a region summary to a second "Summary" sheet (`putRows`, `setCell`,
   `setColumn`), saving to `resources/sales_q1.xlsx` (`saveAs`).
2. Reopens the file (`fromFile`) and inspects it: lists the sheets (`getSheetNames`) and tables
   (`getAllTables`), reads the table's headers and rows (`Table.getHeaders` / `getRows`), finds the
   top order, and spot-checks a single cell (`Sheet.getCell`).

The run prints the workbook structure, the top order, and the spot-checked cell; open
`resources/sales_q1.xlsx` to see the two sheets and the embedded table.
