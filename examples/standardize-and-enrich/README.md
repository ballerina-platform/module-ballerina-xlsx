# Standardize and Enrich

Reads an Excel file in a source-specific layout, maps it onto a standard schema every downstream system
understands, enriches each row (region lookup, computed total, customer tier, and a `time:Date`
processing stamp), and writes the standardised result as a new Excel file. The input is read as a byte
array and parsed with `xlsx:fromBytes`; the output is written with the one-shot `xlsx:writeSheet`.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

1. Creates a self-contained sample source file at `resources/source_sales.xlsx` (short, source-specific
   headers like "Cust" / "Qty").
2. Reads it as bytes (`fileReadBytes`) and parses with `fromBytes` into `SourceRow[]`.
3. Maps each row onto the `StandardRow` schema and enriches it — region from country, line total,
   customer tier, and today's processing date.
4. Writes the standardised, enriched rows to `resources/standardized_sales.xlsx` (`writeSheet`).

The run prints the source and output row counts; compare `resources/source_sales.xlsx` with
`resources/standardized_sales.xlsx` to see the standardisation and enrichment.
