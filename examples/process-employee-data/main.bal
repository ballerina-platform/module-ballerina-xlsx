import ballerina/io;
import ballerinax/xlsx;

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
