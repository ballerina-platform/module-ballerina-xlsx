# Process Employee Data

A Tier 1 quickstart for the one-shot functional API. It writes employee records to a sheet, reads them
back into typed records, filters the engineering team with a Ballerina query expression, and writes the
filtered subset to a new file. Demonstrates `xlsx:writeSheet`, `xlsx:parseSheet`, and `@xlsx:Name`
column-header mapping.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

1. Writes five employee records to `resources/employees.xlsx` (`writeSheet`), with field names mapped to
   the Excel headers "Employee Name", "Department", and "Annual Salary" via `@xlsx:Name`.
2. Reads the file back into `Employee[]` (`parseSheet`).
3. Filters the Engineering rows with a query expression and writes them to `resources/engineers.xlsx`.

The run prints a status line for each file written; open the two files under `resources/` to see the
full and filtered data.
