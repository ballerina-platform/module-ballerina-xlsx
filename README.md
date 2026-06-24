# Ballerina XLSX Library

[![Build](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-timestamped-master.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-timestamped-master.yml)
[![codecov](https://codecov.io/gh/ballerina-platform/module-ballerina-xlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/ballerina-platform/module-ballerina-xlsx)
[![Trivy](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/trivy-scan.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/trivy-scan.yml)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-with-bal-test-graalvm.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-with-bal-test-graalvm.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerina-xlsx.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/commits/main)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/xlsx.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%2Fxlsx)

The Ballerina XLSX library provides functionality to read and write Microsoft Excel files in the XLSX format with type-safe data binding to Ballerina records. It exposes a simple file-based API (`parseSheet` / `writeSheet`) for single-sheet ETL and an object-based Workbook API for multi-sheet operations, byte-array I/O, and Excel Tables.

All processing is done locally with no external service dependencies.

## Quickstart

To use the `xlsx` library in your Ballerina application, modify the `.bal` file as follows:

### Step 1: Import the module

Import the `xlsx` module.

```ballerina
import ballerina/xlsx;
```

### Step 2: Invoke module functions

#### Parse an XLSX file into typed records

```ballerina
type Employee record {|
    string name;
    int age;
    string department;
|};

Employee[] employees = check xlsx:parseSheet("employees.xlsx");
```

#### Write records to an XLSX file

```ballerina
Employee[] employees = [
    {name: "John", age: 30, department: "IT"},
    {name: "Jane", age: 28, department: "HR"}
];

check xlsx:writeSheet(employees, "output.xlsx", "Employees");
```

Writing to an existing file preserves every other sheet — only the named sheet is affected. The write fails by default if that sheet already exists; pass `sheetWriteMode = xlsx:REPLACE` to overwrite it, or `xlsx:APPEND` to add rows below the existing data. The write is atomic — on failure the original file is preserved.

#### Map non-matching headers with `@xlsx:Name`

```ballerina
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};

Employee[] employees = check xlsx:parseSheet("employees.xlsx");
```

#### Work with multiple sheets via the Workbook API

```ballerina
xlsx:Workbook wb = check xlsx:fromFile("report.xlsx");

xlsx:Sheet sales = check wb.getSheet("Sales");
Employee[] salesRows = check sales.getRows();

xlsx:Sheet summary = check wb.createSheet("Summary");
check summary.putRows(salesRows);

check wb.save();
check wb.close();
```

#### Read and write Excel Tables

For one-shot table-by-name flows, the tier 1 functions are simplest (tables are unique across the workbook, so no sheet specifier is needed):

```ballerina
Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");

Employee[] additions = [{name: "Alice", age: 31, department: "Eng"}];
check xlsx:writeTable([...employees, ...additions], "sales.xlsx", "EmployeeTable");
// writeTable resizes the table's data range to fit the data (grows or shrinks)
```

For totals rows, rename, resize, or coordination with other workbook operations, go through the Workbook API:

```ballerina
xlsx:Workbook wb = check xlsx:fromFile("sales.xlsx");
xlsx:Table empTable = check wb.getTable("EmployeeTable");

Employee[] employees = check empTable.getRows();
if check empTable.hasTotalRow() {
    map<xlsx:CellValue> totals = check empTable.getTotalRow();
    // ...
}

check wb.save();
check wb.close();
```

#### Bind dates and times to `time:Civil`, `time:Date`, or `time:TimeOfDay`

Declare the field's type to control the shape — typed time records, or ISO 8601 strings:

```ballerina
import ballerina/time;

type Transaction record {|
    int id;
    time:Civil timestamp;      // date-time cell → time:Civil
    time:Date settledOn;       // date-only cell → time:Date
    decimal amount;
|};

Transaction[] txns = check xlsx:parseSheet("transactions.xlsx");
```

#### Read from bytes, write to bytes

```ballerina
byte[] inputBytes = check sftp->get("/in/report.xlsx");
xlsx:Workbook wb = check xlsx:fromBytes(inputBytes);

// ...read, modify...

byte[] outputBytes = check wb.toBytes();
check sftp->put("/out/report.xlsx", outputBytes);
check wb.close();
```

### Step 3: Run the Ballerina application

```bash
bal run
```

## Examples

The `xlsx` library provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/), progressing from the simplest Tier 1 read/write loop through multi-sheet workbooks, validation, and in-memory byte pipelines to database and enrichment flows.

1. [Process Employee Data](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/process-employee-data) — Tier 1 quickstart. Write employee records, read them back into typed records, filter, write the filtered subset. Demonstrates `parseSheet`, `writeSheet`, and `@xlsx:Name` column mapping.
2. [Monthly Sales Report](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/monthly-sales-report) — Build a multi-sheet workbook with an embedded Excel Table and `time:Date` columns; reopen and query it through the Workbook + Table APIs.
3. [Validated Bulk Import](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/validated-bulk-import) — Parse a partner file with `@constraint` validation and fail-safe error logging — clean rows flow downstream; rejected rows are logged with their raw values and reason.
4. [In-Memory Pipeline](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/in-memory-pipeline) — Process XLSX bytes end-to-end without disk I/O. Demonstrates `xlsx:fromBytes` and `Workbook.toBytes()` — the shape an HTTP service or queue consumer would use.
5. [Database to Excel](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/database-to-excel) — Read rows from a database (in-memory H2, no server to configure), map them onto a consumer's column layout with `@xlsx:Name`, build the workbook with the Workbook API, and serialise it to bytes with `Workbook.toBytes()`.
6. [Standardize and Enrich](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/standardize-and-enrich) — Parse an Excel file's bytes with `xlsx:fromBytes`, map its layout onto a standard schema, enrich each row (region lookup, computed total, customer tier, `time:Date` stamp), and write the result with `writeSheet`.

## Issues and projects

Issues and Projects tabs are disabled for this repository as this is part of the Ballerina library. To report bugs, request new features, start new discussions, view project boards, etc., visit the Ballerina library [parent repository](https://github.com/ballerina-platform/ballerina-library).

This repository only contains the source code for the package.

## Build from the source

### Setting up the prerequisites

1. Download and install Java SE Development Kit (JDK) version 21. You can download it from either of the following sources:

    * [Oracle JDK](https://www.oracle.com/java/technologies/downloads/)
    * [OpenJDK](https://adoptium.net/)

   > **Note:** After installation, remember to set the `JAVA_HOME` environment variable to the directory where JDK was installed.

2. Download and install [Ballerina Swan Lake](https://ballerina.io/).

3. Download and install [Docker](https://www.docker.com/get-started).

   > **Note**: Ensure that the Docker daemon is running before executing any tests.

4. Export Github Personal access token with read package permissions as follows,

    ```bash
    export packageUser=<Username>
    export packagePAT=<Personal access token>
    ```

### Build options

Execute the commands below to build from the source.

1. To build the package:

   ```bash
   ./gradlew clean build
   ```

2. To run the tests:

   ```bash
   ./gradlew clean test
   ```

3. To build the without the tests:

   ```bash
   ./gradlew clean build -x test
   ```

4. To run tests against different environments:

   ```bash
   ./gradlew clean test -Pgroups=<Comma separated groups/test cases>
   ```

5. To debug the package with a remote debugger:

   ```bash
   ./gradlew clean build -Pdebug=<port>
   ```

6. To debug with the Ballerina language:

   ```bash
   ./gradlew clean build -PbalJavaDebug=<port>
   ```

7. Publish the generated artifacts to the local Ballerina Central repository:

    ```bash
    ./gradlew clean build -PpublishToLocalCentral=true
    ```

8. Publish the generated artifacts to the Ballerina Central repository:

   ```bash
   ./gradlew clean build -PpublishToCentral=true
   ```

## Contribute to Ballerina

As an open-source project, Ballerina welcomes contributions from the community.

For more information, go to the [contribution guidelines](https://github.com/ballerina-platform/ballerina-lang/blob/master/CONTRIBUTING.md).

## Code of conduct

All the contributors are encouraged to read the [Ballerina Code of Conduct](https://ballerina.io/code-of-conduct).

## Useful links

* For more information go to the [`xlsx` package](https://central.ballerina.io/ballerina/xlsx/latest).
* For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
