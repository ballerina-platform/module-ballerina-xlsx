/*
 * Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
 *
 * WSO2 LLC. licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package io.ballerina.stdlib.xlsx.utils;

import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BString;

/**
 * Constants used throughout the XLSX module.
 */
public final class XlsxConstants {

    private XlsxConstants() {
        // Private constructor to prevent instantiation
    }

    // Parse options field names
    public static final BString HEADER_ROW_INDEX = StringUtils.fromString("headerRowIndex");
    public static final BString DATA_START_ROW_INDEX = StringUtils.fromString("dataStartRowIndex");
    public static final BString ROW_COUNT = StringUtils.fromString("rowCount");
    public static final BString FORMULA_MODE = StringUtils.fromString("formulaMode");
    public static final BString ENABLE_CONSTRAINT_VALIDATION = StringUtils.fromString("enableConstraintValidation");
    public static final BString CASE_INSENSITIVE_HEADERS = StringUtils.fromString("caseInsensitiveHeaders");

    // Data projection options
    public static final BString ALLOW_DATA_PROJECTION = StringUtils.fromString("allowDataProjection");
    public static final BString NIL_AS_OPTIONAL_FIELD = StringUtils.fromString("nilAsOptionalField");
    public static final BString ABSENT_AS_NILABLE_TYPE = StringUtils.fromString("absentAsNilableType");

    // Write options field names
    public static final BString WRITE_SHEET_NAME = StringUtils.fromString("sheetName");
    public static final BString WRITE_HEADERS = StringUtils.fromString("writeHeaders");
    public static final BString START_ROW_INDEX = StringUtils.fromString("startRowIndex");
    // Internal-only — not exposed on public WriteOptions. Used by SheetHandle.createTableFromData
    // to plumb its startColumnIndex through to XlsxWriter without broadening the public API.
    public static final BString START_COLUMN_INDEX = StringUtils.fromString("startColumnIndex");

    // Formula mode values
    public static final String FORMULA_MODE_CACHED = "CACHED";
    public static final String FORMULA_MODE_TEXT = "TEXT";

    // Ballerina type names — Sheet and Table are public `object` types in Ballerina
    // with non-public `SheetImpl` / `TableImpl` classes providing the implementation.
    // Native instance construction must target the concrete class name.
    public static final String WORKBOOK_TYPE = "Workbook";
    public static final String SHEET_TYPE = "SheetImpl";
    public static final String TABLE_TYPE = "TableImpl";

    // Default values
    public static final String DEFAULT_SHEET_NAME = "Sheet1";
    public static final int DEFAULT_HEADER_ROW = 0;

    // Error type names
    public static final String ERROR_TYPE = "Error";
    public static final String PARSE_ERROR_TYPE = "ParseError";
    public static final String FILE_NOT_FOUND_ERROR_TYPE = "FileNotFoundError";
    public static final String SHEET_NOT_FOUND_ERROR_TYPE = "SheetNotFoundError";
    public static final String TYPE_CONVERSION_ERROR_TYPE = "TypeConversionError";
    public static final String CONSTRAINT_VALIDATION_ERROR_TYPE = "ConstraintValidationError";
    public static final String TABLE_NOT_FOUND_ERROR_TYPE = "TableNotFoundError";
    public static final String TABLE_OVERLAP_ERROR_TYPE = "TableOverlapError";
    public static final String INVALID_TABLE_RANGE_ERROR_TYPE = "InvalidTableRangeError";

    // Limits
    public static final int MAX_FILE_SIZE_MB = 100;
    public static final int MAX_ROWS = 1_048_576;  // Excel max rows
    public static final int MAX_COLUMNS = 16_384;   // Excel max columns

    // Fail-safe options field names
    public static final BString FAIL_SAFE = StringUtils.fromString("failSafe");
    public static final BString ENABLE_CONSOLE_LOGS = StringUtils.fromString("enableConsoleLogs");
    public static final BString INCLUDE_SOURCE_DATA_IN_CONSOLE = StringUtils.fromString("includeSourceDataInConsole");
    public static final BString FILE_OUTPUT_MODE = StringUtils.fromString("fileOutputMode");
    public static final BString FILE_PATH = StringUtils.fromString("filePath");
    public static final BString CONTENT_TYPE = StringUtils.fromString("contentType");
    public static final BString FILE_WRITE_OPTION = StringUtils.fromString("fileWriteOption");
    public static final BString OFFENDING_ROW = StringUtils.fromString("offendingRow");

    // Fail-safe content type values
    public static final String CONTENT_TYPE_METADATA = "METADATA";
    public static final String CONTENT_TYPE_RAW = "RAW";
    public static final String CONTENT_TYPE_RAW_AND_METADATA = "RAW_AND_METADATA";

    // Fail-safe file write option values
    public static final String FILE_WRITE_APPEND = "APPEND";
    public static final String FILE_WRITE_OVERWRITE = "OVERWRITE";

    // Fail-safe function name
    public static final String PRINT_ERROR = "printError";

    // Fail-safe warning messages
    public static final String XLSX_PARSE_ERROR = "XLSX parse warning at row %d, column %d: %s";
    public static final String FILE_IO_ERROR = "Failed to create log file at: %s. Caused by: %s";
    public static final String FILE_OVERWRITE_ERROR = "Failed to overwrite log file at: %s. Caused by: %s";
    public static final String FILE_WRITE_ERROR = "Failed to write log file at: %s. Caused by: %s";
}
