# Examples

The `ballerina/xlsx` library provides practical examples illustrating usage in various scenarios. The set below progresses from the simplest Tier 1 read/write loop to in-memory byte pipelines suitable for an HTTP handler — reading them in order forms a complete picture of the module's surface.

1. [Process Employee Data](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/process-employee-data) - Tier 1 quickstart. Write employee records, read them back into typed records, filter, write the filtered subset to a new file. Demonstrates `parseSheet`, `writeSheet`, and `@xlsx:Name` column mapping.
2. [Monthly Sales Report](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/monthly-sales-report) - Build a multi-sheet workbook from scratch with an embedded Excel Table and `time:Date` columns; reopen the file and query it back through the Workbook + Table APIs.
3. [Validated Bulk Import](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/validated-bulk-import) - Parse a partner-supplied file with `@constraint` validation and fail-safe error logging — clean rows flow downstream; rejected rows are logged with their raw values and the reason for rejection.
4. [In-Memory Pipeline](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/in-memory-pipeline) - Process XLSX bytes end-to-end without disk I/O. Demonstrates `xlsx:fromBytes` and `Workbook.toBytes()` — the shape an HTTP service or queue consumer would use.
5. [Database to Excel](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/database-to-excel) - Read rows from a database (in-memory H2, no server to configure), map them onto a consumer's column layout with `@xlsx:Name`, build the workbook with the Workbook API, serialise it to a byte array with `Workbook.toBytes()`, and write the bytes to a file.
6. [Standardize and Enrich](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/standardize-and-enrich) - Read an Excel file's bytes and parse them with `xlsx:fromBytes`, map its layout onto a standard schema, enrich each row (region lookup, computed total, customer tier, `time:Date` processing stamp), and write the result with `writeSheet`.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later

## Running an example

Execute the following commands to build an example from the source:

* To build an example:

    ```bash
    bal build
    ```

* To run an example:

    ```bash
    bal run
    ```

## Building the examples with the local module

**Warning**: Due to the absence of support for reading local repositories for single Ballerina files, the Bala of the module is manually written to the central repository as a workaround. Consequently, the bash script may modify your local Ballerina repositories.

Execute the following commands to build all the examples against the changes you have made to the module locally:

* To build all the examples:

    ```bash
    ./build.sh build
    ```

* To run all the examples:

    ```bash
    ./build.sh run
    ```
