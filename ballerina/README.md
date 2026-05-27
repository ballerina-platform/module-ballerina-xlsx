## Overview

The `ballerina/xlsx` module provides functionality to read and write Microsoft Excel files in the XLSX format with type-safe data binding to Ballerina records. It exposes a simple file-based API (`parseSheet` / `writeSheet`) for single-sheet ETL and an object-based Workbook API for multi-sheet operations, byte-array I/O, and Excel Tables.

All processing is done locally with no external service dependencies.

## Quickstart

To use the `xlsx` module in your Ballerina application, modify the `.bal` file as follows:

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

You can also target a specific sheet by name or index, or fall back to a flexible shape:

```ballerina
Employee[] sales = check xlsx:parseSheet("report.xlsx", "Sales");
map<anydata>[] rows = check xlsx:parseSheet("unknown.xlsx");
string[][] raw = check xlsx:parseSheet("anything.xlsx");
```

#### Write records to an XLSX file

```ballerina
Employee[] employees = [
    {name: "John", age: 30, department: "IT"},
    {name: "Jane", age: 28, department: "HR"}
];

check xlsx:writeSheet(employees, "output.xlsx", sheetName = "Employees");
```

The write is atomic — on failure the original file is preserved, never partially overwritten.

#### Map non-matching headers with `@xlsx:Name`

When the Excel column header does not match the Ballerina field name, use the `@xlsx:Name` annotation. The mapping is bidirectional — it applies on both parse and write.

```ballerina
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};

Employee[] employees = check xlsx:parseSheet("employees.xlsx");
```

#### Work with multiple sheets via the `Workbook` API

```ballerina
xlsx:Workbook wb = check new("report.xlsx");

string[] sheetNames = wb.getSheetNames();
xlsx:Sheet sales = check wb.getSheet("Sales");
Employee[] salesRows = check sales.getRows();

xlsx:Sheet summary = check wb.createSheet("Summary");
check summary.putRows(salesRows);

check wb.save();
check wb.close();
```

The same `Workbook` constructor accepts three forms:

```ballerina
xlsx:Workbook wb1 = check new;                       // empty in-memory workbook (saveAs required to persist)
xlsx:Workbook wb2 = check new("existing.xlsx");      // open an existing file (errors if missing)
xlsx:Workbook wb3 = check new(sourceBytes);          // open from a byte array (e.g., SFTP / HTTP body)
```

To create a brand-new file with a specific name, use `new` (in-memory) followed by `wb.saveAs("path.xlsx")`.

#### Bytes in, bytes out

```ballerina
byte[] inputBytes = check sftp->get("/in/orders.xlsx");
xlsx:Workbook wb = check new(inputBytes);

xlsx:Sheet sheet = check wb.getSheet(0);
Order[] orders = check sheet.getRows();
// ... enrich orders ...
check sheet.putRows(orders);

byte[] outputBytes = check wb.toBytes();
check sftp->put("/out/orders-enriched.xlsx", outputBytes);
check wb.close();
```

#### Read and write Excel Tables

Excel Tables (ListObjects) are addressed by name and are unique across the workbook. For one-shot read/write flows, use the tier 1 conveniences:

```ballerina
Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");

Employee[] newEmployees = [{name: "Alice", age: 31, department: "Eng"}];
check xlsx:writeTable([...employees, ...newEmployees], "sales.xlsx", "EmployeeTable");
// writeTable auto-expands the table to fit the data
```

For richer operations (totals row, rename, resize, or coordination with other workbook changes), go through the Workbook API:

```ballerina
xlsx:Workbook wb = check new("sales.xlsx");
xlsx:Table empTable = check wb.getTable("EmployeeTable");

Employee[] employees = check empTable.getRows();
if empTable.hasTotalsRow() {
    map<anydata> totals = check empTable.getTotalsRow();
    // ...
}
check empTable.putRows([...employees, ...newEmployees]);

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

Writing the same record back produces date-formatted cells, not text cells. Use `string` instead of the `time:*` types if you want ISO strings in your record.

#### Continue parsing on row-level errors with fail-safe mode

```ballerina
Employee[] employees = check xlsx:parseSheet("data.xlsx", 0, {
    failSafe: {
        enableConsoleLogs: true,
        fileOutputMode: {
            filePath: "./xlsx-errors.log",
            contentType: RAW_AND_METADATA
        }
    }
});
```

Rows that fail type conversion or constraint validation are skipped and logged; structural errors (corrupted file, missing sheet) still fail immediately.

### Step 3: Run the Ballerina application

```bash
bal run
```

## Examples

The `xlsx` module provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/).

1. [Process employee data](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/process-employee-data/) — Reads employee records from an XLSX file, processes them, and writes the results to a new file.
