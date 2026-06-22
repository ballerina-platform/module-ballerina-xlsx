# Validated Bulk Import

A downstream pipeline accepting a partner's XLSX upload. Real-world data is messy — some rows fail
business rules — so this keeps the pipeline clean while preserving full visibility into rejected rows.
It applies Ballerina `@constraint` annotations to the target record, parses with
`enableConstraintValidation` plus `FailSafeOptions` (file logging), and uses `caseInsensitiveHeaders`
to tolerate partner-side header capitalisation. Valid rows flow through; invalid rows are skipped and
logged with their raw cell values and the reason for rejection.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

1. Writes a simulated partner file to `resources/customers.xlsx` (several rows deliberately violate the
   `@constraint` rules, and headers use mixed capitalisation).
2. Parses it with constraint validation and fail-safe file logging
   (`contentType: RAW_AND_METADATA`, `fileWriteOption: OVERWRITE`). Clean rows are returned to the
   caller; rejected rows are written to `resources/import_errors.json`.

The run prints the accepted rows and the contents of the error log. **What to look for:**
`resources/import_errors.json` — each rejected row with its raw values and the validation failure.
