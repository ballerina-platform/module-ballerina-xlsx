# In-Memory Pipeline

A microservice receives an XLSX as a byte array (HTTP upload, queue payload, object-storage download —
any byte source), applies a transformation, and returns the modified workbook as a byte array, with no
disk I/O at any point. This is the primary use case for the byte-array API surface (`xlsx:fromBytes` /
`Workbook.toBytes()`). The handler here is a pure `byte[] -> byte[]|error` function — the exact shape
that drops into an HTTP handler, queue consumer, or batch worker unchanged.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Run

```bash
bal run
```

> Before `ballerina/xlsx` is published to Ballerina Central, build the examples against the local module
> using the script described in the [examples README](../README.md#building-the-examples-with-the-local-module).

## What it does

Everything runs in memory — there is no `resources/` directory and nothing is written to disk.

1. Builds an "incoming" invoice workbook and serialises it to bytes (`toBytes`), simulating a payload
   arriving over the wire.
2. Runs the handler `applyTaxSurcharge(byte[], rate)`: loads the bytes (`fromBytes`), discovers the
   data extent (`getUsedCellRange`), applies a 10% surcharge to the `amount` column cell-by-cell
   (`getCell` / `setCell`), records an audit note on a separate sheet (`setCellByAddress`), and returns
   the modified workbook as bytes (`toBytes`).
3. Loads the outgoing bytes back (`fromBytes`) and prints the transformed invoices to verify.
