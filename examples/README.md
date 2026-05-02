# Examples

The `ballerinax/xlsx` library provides practical examples illustrating usage in various scenarios.

1. [Process Employee Data](https://github.com/ballerina-platform/module-ballerina-xlsx/tree/main/examples/process-employee-data) - Write employee records to an XLSX file, read them back into typed records, filter by department, and write the results to a new file.

## Prerequisites

- Ballerina Swan Lake 2201.13.1 or later

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
