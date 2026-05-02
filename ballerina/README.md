## Overview

The `ballerinax/xlsx` module provides functionality to read and write Microsoft Excel files in the XLSX format with type-safe data binding to Ballerina records. It exposes a simple file-based API for single-sheet operations and a richer Workbook/Sheet/Table API for multi-sheet workbooks and Excel Tables.

All processing is done locally with no external service dependencies.

## Quickstart

To use the `xlsx` module in your Ballerina application, modify the `.bal` file as follows:

### Step 1: Import the module

Import the `xlsx` module.

```ballerina
import ballerinax/xlsx;
```

### Step 2: Invoke module functions

#### Parse an XLSX file into typed records

```ballerina
type Employee record {|
    string name;
    int age;
    string department;
|};

Employee[] employees = check xlsx:parse("employees.xlsx");
```

You can also target a specific sheet by name or index, or fall back to a flexible shape:

```ballerina
Employee[] sales = check xlsx:parse("report.xlsx", "Sales");
map<anydata>[] rows = check xlsx:parse("unknown.xlsx");
string[][] raw = check xlsx:parse("anything.xlsx");
```

#### Write records to an XLSX file

```ballerina
Employee[] employees = [
    {name: "John", age: 30, department: "IT"},
    {name: "Jane", age: 28, department: "HR"}
];

check xlsx:write(employees, "output.xlsx", sheetName = "Employees");
```

#### Map non-matching headers with `@xlsx:Name`

When the Excel column header does not match the Ballerina field name, use the `@xlsx:Name` annotation. The mapping is bidirectional — it applies on both parse and write.

```ballerina
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};

Employee[] employees = check xlsx:parse("employees.xlsx");
```

#### Work with multiple sheets via the `Workbook` API

```ballerina
xlsx:Workbook wb = check xlsx:openFile("report.xlsx");

string[] sheetNames = wb.getSheetNames();
xlsx:Sheet sales = check wb.getSheet("Sales");
Employee[] salesRows = check sales.getRows();

xlsx:Sheet summary = check wb.createSheet("Summary");
check summary.putRows(salesRows);

check wb.save();
check wb.close();
```

To create a brand-new file, use `xlsx:createFile(path)` (writes on `save()`) or `xlsx:createWorkbook()` (in-memory; requires `saveAs(path)`).

#### Read and write Excel Tables

Excel Tables (ListObjects) are addressed by name and are unique across the workbook.

```ballerina
Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");

Employee[] newEmployees = [{name: "Alice", age: 31, department: "Eng"}];
check xlsx:writeTable(newEmployees, "sales.xlsx", "EmployeeTable");
```

The Table object also supports headers, totals row, rename, and resize operations via `wb.getTable("EmployeeTable")`.

#### Continue parsing on row-level errors with fail-safe mode

```ballerina
Employee[] employees = check xlsx:parse("data.xlsx", 0, {
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
