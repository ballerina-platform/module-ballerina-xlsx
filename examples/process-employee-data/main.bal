// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/io;
import ballerina/xlsx;

// Map Excel column headers to record fields
type Employee record {|
    @xlsx:Name {value: "Employee Name"}
    string name;
    @xlsx:Name {value: "Department"}
    string department;
    @xlsx:Name {value: "Annual Salary"}
    decimal salary;
|};

public function main() returns error? {
    // Create sample employee data
    Employee[] employees = [
        {name: "Alice", department: "Engineering", salary: 95000},
        {name: "Bob", department: "Marketing", salary: 72000},
        {name: "Carol", department: "Engineering", salary: 88000},
        {name: "Dave", department: "Sales", salary: 67000},
        {name: "Eve", department: "Engineering", salary: 102000}
    ];
    check xlsx:writeSheet(employees, "resources/employees.xlsx");
    io:println(string `Created employees.xlsx`);

    // Read the spreadsheet back into typed records
    Employee[] parsed = check xlsx:parseSheet("resources/employees.xlsx");

    // Filter engineering team and write to a new file
    Employee[] engineers = from Employee e in parsed
        where e.department == "Engineering"
        select e;

    check xlsx:writeSheet(engineers, "resources/engineers.xlsx");
    io:println(string `Wrote ${engineers.length()} engineers to engineers.xlsx`);
}
