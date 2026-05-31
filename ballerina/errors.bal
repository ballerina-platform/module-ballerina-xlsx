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

# Represents a generic XLSX module error.
public type Error distinct error<ErrorDetails>;

# Represents an error that occurs during XLSX parsing.
public type ParseError distinct Error;

# Represents an error when the specified file cannot be found or accessed.
public type FileNotFoundError distinct Error;

# Represents an error when a requested sheet is not found.
public type SheetNotFoundError distinct Error;

# Represents an error during type conversion.
public type TypeConversionError distinct Error;

# Represents an error when constraint validation fails for a record field.
public type ConstraintValidationError distinct Error;

# Represents an error when a requested table is not found.
public type TableNotFoundError distinct Error;

# Represents an error when creating a table would overlap with an existing table.
public type TableOverlapError distinct Error;

# Represents an error when a table range specification is invalid.
public type InvalidTableRangeError distinct Error;

# Details for XLSX errors.
public type ErrorDetails record {|
    # Name of the sheet where error occurred (if applicable)
    string sheetName?;
    # Name of the table where error occurred (if applicable)
    string tableName?;
    # Cell address where error occurred (if applicable)
    string cellAddress?;
    # Row number where error occurred (if applicable)
    int rowNumber?;
    # Column number where error occurred (if applicable)
    int columnNumber?;
|};
