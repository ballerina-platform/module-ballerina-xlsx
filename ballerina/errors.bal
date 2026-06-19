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

# The base type for all `xlsx` module errors.
public type Error distinct error<ErrorDetails>;

# The workbook content is malformed or could not be read.
public type ParseError distinct Error;

# The XLSX file path does not exist or could not be accessed.
public type FileNotFoundError distinct Error;

# No sheet matches the given name or index.
public type SheetNotFoundError distinct Error;

# A sheet with the target name already exists.
public type SheetExistsError distinct Error;

# A cell value could not be converted to the target type.
public type TypeConversionError distinct Error;

# A parsed record failed a `@constraint` rule.
public type ConstraintValidationError distinct Error;

# No table matches the given name.
public type TableNotFoundError distinct Error;

# A table with the target name already exists.
public type TableExistsError distinct Error;

# A table write would overlap with another table.
public type TableOverlapError distinct Error;

# A table range or insert position is invalid.
public type InvalidTableRangeError distinct Error;

# Details attached to an XLSX error.
public type ErrorDetails record {|
    # Sheet where the error occurred
    string sheetName?;
    # Table where the error occurred
    string tableName?;
    # Cell address where the error occurred
    string cellAddress?;
    # Row number where the error occurred
    int rowNumber?;
    # Column number where the error occurred
    int columnNumber?;
    # Record field involved in the error
    string fieldName?;
|};
