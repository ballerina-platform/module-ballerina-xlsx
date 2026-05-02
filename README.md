# Ballerina XLSX Library

[![Build](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-timestamped-master.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-timestamped-master.yml)
[![codecov](https://codecov.io/gh/ballerina-platform/module-ballerina-xlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/ballerina-platform/module-ballerina-xlsx)
[![Trivy](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/trivy-scan.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/trivy-scan.yml)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-with-bal-test-graalvm.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/actions/workflows/build-with-bal-test-graalvm.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerina-xlsx.svg)](https://github.com/ballerina-platform/module-ballerina-xlsx/commits/main)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/xlsx.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%2Fxlsx)

The `ballerinax/xlsx` library provides functionality to read and write Microsoft Excel files in the XLSX format with type-safe data binding to Ballerina records. It exposes a simple file-based API for single-sheet operations and a richer Workbook/Sheet/Table API for multi-sheet workbooks and Excel Tables.

All processing is done locally with no external service dependencies.

## Quickstart

To use the `xlsx` library in your Ballerina application, modify the `.bal` file as follows:

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

#### Write records to an XLSX file

```ballerina
Employee[] employees = [
    {name: "John", age: 30, department: "IT"},
    {name: "Jane", age: 28, department: "HR"}
];

check xlsx:write(employees, "output.xlsx", sheetName = "Employees");
```

#### Map non-matching headers with `@xlsx:Name`

```ballerina
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;

    @xlsx:Name {value: "Years of Service"}
    int tenure;
|};

Employee[] employees = check xlsx:parse("employees.xlsx");
```

#### Work with multiple sheets via the Workbook API

```ballerina
xlsx:Workbook wb = check xlsx:openFile("report.xlsx");

xlsx:Sheet sales = check wb.getSheet("Sales");
Employee[] salesRows = check sales.getRows();

xlsx:Sheet summary = check wb.createSheet("Summary");
check summary.putRows(salesRows);

check wb.save();
check wb.close();
```

#### Read and write Excel Tables

```ballerina
Employee[] employees = check xlsx:parseTable("sales.xlsx", "EmployeeTable");

Employee[] additions = [{name: "Alice", age: 31, department: "Eng"}];
check xlsx:writeTable(additions, "sales.xlsx", "EmployeeTable");
```

### Step 3: Run the Ballerina application

```bash
bal run
```

## Examples

The `xlsx` library provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/).

1. [Process employee data](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/process-employee-data/) — Reads employee records from an XLSX file, processes them, and writes the results to a new file.

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

* For more information go to the [`xlsx` package](https://central.ballerina.io/ballerinax/xlsx/latest).
* For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
