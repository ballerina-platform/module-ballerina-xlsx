# Database to Excel

Exports rows from a database to an Excel file. It reads the source rows, maps them onto the column
layout a downstream consumer expects with `@xlsx:Name` (deriving a "Line Total" column along the way),
builds the workbook with the Workbook API, serialises it with `Workbook.toBytes()`, and writes the
bytes to a file.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later
- No external database needed — the example seeds an **in-memory H2** instance on the fly
  (`jdbc:h2:mem:`, pulled in via `ballerinax/java.jdbc` + `ballerinax/h2.driver`). Point `jdbc:Client`
  at your own database to use real data.

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

1. Seeds an in-memory H2 database with sample orders and reads them back (`jdbc:Client`).
2. Maps each database row onto the consumer's `VendorOrder` schema, computing the line total, and writes
   them to a sheet (`putRows`).
3. Serialises the workbook to bytes (`toBytes`) and writes them to `resources/orders.xlsx`.

The run prints the row and byte counts; open `resources/orders.xlsx` to see the exported layout.
