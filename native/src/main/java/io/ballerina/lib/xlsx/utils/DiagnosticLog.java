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

package io.ballerina.lib.xlsx.utils;

import io.ballerina.runtime.api.creators.ErrorCreator;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

import java.text.MessageFormat;
import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Utility class for creating diagnostic errors.
 */
public final class DiagnosticLog {

    private static final String ERROR_PREFIX = "error.";
    private static ResourceBundle errorBundle;

    // Ballerina error type names (as declared in errors.bal).
    private static final String ERROR_TYPE = "Error";
    private static final String PARSE_ERROR_TYPE = "ParseError";
    private static final String FILE_NOT_FOUND_ERROR_TYPE = "FileNotFoundError";
    private static final String SHEET_NOT_FOUND_ERROR_TYPE = "SheetNotFoundError";
    private static final String TYPE_CONVERSION_ERROR_TYPE = "TypeConversionError";
    private static final String CONSTRAINT_VALIDATION_ERROR_TYPE = "ConstraintValidationError";
    private static final String TABLE_NOT_FOUND_ERROR_TYPE = "TableNotFoundError";
    private static final String TABLE_OVERLAP_ERROR_TYPE = "TableOverlapError";
    private static final String INVALID_TABLE_RANGE_ERROR_TYPE = "InvalidTableRangeError";

    private DiagnosticLog() {
        // Private constructor to prevent instantiation
    }

    static {
        try {
            errorBundle = ResourceBundle.getBundle("xlsx_error", Locale.getDefault());
        } catch (Exception e) {
            // Fallback: errors will use default messages
            errorBundle = null;
        }
    }

    /**
     * Create a generic XLSX error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError error(String message) {
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                ERROR_TYPE,
                StringUtils.fromString(message),
                null,
                null
        );
    }

    /**
     * Create a generic XLSX error with cause.
     *
     * @param message Error message
     * @param cause   Cause of the error
     * @return BError
     */
    public static BError error(String message, Throwable cause) {
        BError causeError = cause instanceof BError ? (BError) cause :
                ErrorCreator.createError(StringUtils.fromString(cause.getMessage()));
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                ERROR_TYPE,
                StringUtils.fromString(message),
                causeError,
                null
        );
    }

    /**
     * Create a diagnostic error with code and arguments.
     *
     * @param code Error code
     * @param args Message arguments
     * @return BError
     */
    public static BError error(DiagnosticErrorCode code, Object... args) {
        String message = formatMessage(code, args);
        return error(message);
    }

    /**
     * Create a parse error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError parseError(String message) {
        return createTypedError(PARSE_ERROR_TYPE, message, null);
    }

    /**
     * Create a parse error with details.
     *
     * @param message   Error message
     * @param sheetName Sheet name where error occurred
     * @param row       Row number where error occurred
     * @param column    Column number where error occurred
     * @return BError
     */
    public static BError parseError(String message, String sheetName, Integer row, Integer column) {
        BMap<BString, Object> details = createErrorDetails(sheetName, null, null, row, column);
        return createTypedError(PARSE_ERROR_TYPE, message, details);
    }

    /**
     * Create a file not found error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError fileNotFoundError(String message) {
        return createTypedError(FILE_NOT_FOUND_ERROR_TYPE, message, null);
    }

    /**
     * Create a file not found error with cause.
     *
     * @param message Error message
     * @param cause   Cause of the error
     * @return BError
     */
    public static BError fileNotFoundError(String message, Throwable cause) {
        BError causeError = cause instanceof BError ? (BError) cause :
                ErrorCreator.createError(StringUtils.fromString(cause.getMessage()));
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                FILE_NOT_FOUND_ERROR_TYPE,
                StringUtils.fromString(message),
                causeError,
                null
        );
    }

    /**
     * Create a sheet not found error.
     *
     * @param sheetName Sheet name that was not found
     * @return BError
     */
    public static BError sheetNotFoundError(String sheetName) {
        String message = "Sheet '" + sheetName + "' not found in workbook";
        BMap<BString, Object> details = createErrorDetails(sheetName, null, null, null, null);
        return createTypedError(SHEET_NOT_FOUND_ERROR_TYPE, message, details);
    }

    /**
     * Create a sheet not found error for index-based access.
     *
     * @param index Sheet index that was not found
     * @param maxIndex Maximum valid index
     * @return BError
     */
    public static BError sheetNotFoundError(int index, int maxIndex) {
        String message = "Sheet at index " + index + " not found (valid range: 0-" + maxIndex + ")";
        return createTypedError(SHEET_NOT_FOUND_ERROR_TYPE, message, null);
    }

    /**
     * Create a type conversion error.
     *
     * @param message     Error message
     * @param cellAddress Cell address where error occurred
     * @param row         Row number
     * @param column      Column number
     * @return BError
     */
    public static BError typeConversionError(String message, String cellAddress, Integer row, Integer column) {
        BMap<BString, Object> details = createErrorDetails(null, null, cellAddress, row, column);
        return createTypedError(TYPE_CONVERSION_ERROR_TYPE, message, details);
    }

    /**
     * Create a constraint validation error.
     *
     * @param message Error message from constraint validation
     * @param row     Row number where validation failed
     * @param column  Column number where validation failed (if applicable)
     * @return BError
     */
    public static BError constraintValidationError(String message, Integer row, Integer column) {
        BMap<BString, Object> details = createErrorDetails(null, null, null, row, column);
        return createTypedError(CONSTRAINT_VALIDATION_ERROR_TYPE, message, details);
    }

    /**
     * Create a table not found error.
     *
     * @param tableName Table name that was not found
     * @return BError
     */
    public static BError tableNotFoundError(String tableName) {
        String message = "Table '" + tableName + "' not found in workbook";
        BMap<BString, Object> details = createErrorDetails(null, tableName, null, null, null);
        return createTypedError(TABLE_NOT_FOUND_ERROR_TYPE, message, details);
    }

    /**
     * Create a table not found error with sheet context.
     *
     * @param tableName Table name that was not found
     * @param sheetName Sheet name where table was expected
     * @return BError
     */
    public static BError tableNotFoundError(String tableName, String sheetName) {
        String message = "Table '" + tableName + "' not found in sheet '" + sheetName + "'";
        BMap<BString, Object> details = createErrorDetails(sheetName, tableName, null, null, null);
        return createTypedError(TABLE_NOT_FOUND_ERROR_TYPE, message, details);
    }

    /**
     * Create a table overlap error.
     *
     * @param tableName     Name of the table being created
     * @param existingTable Name of the existing overlapping table
     * @param sheetName     Sheet name where overlap occurs
     * @return BError
     */
    public static BError tableOverlapError(String tableName, String existingTable, String sheetName) {
        String message = "Cannot create table '" + tableName + "': range overlaps with existing table '" +
                existingTable + "' in sheet '" + sheetName + "'";
        BMap<BString, Object> details = createErrorDetails(sheetName, tableName, null, null, null);
        return createTypedError(TABLE_OVERLAP_ERROR_TYPE, message, details);
    }

    /**
     * Create a table-overlap error for a write that would resize a table over another table.
     *
     * @param tableName     Name of the table being written
     * @param existingTable Name of the table the resize would collide with
     * @param sheetName     Sheet name where the collision occurs
     * @return BError
     */
    public static BError tableResizeOverlapError(String tableName, String existingTable, String sheetName) {
        String message = "Cannot write to table '" + tableName + "': resizing would overlap table '" +
                existingTable + "' in sheet '" + sheetName + "'";
        BMap<BString, Object> details = createErrorDetails(sheetName, tableName, null, null, null);
        return createTypedError(TABLE_OVERLAP_ERROR_TYPE, message, details);
    }

    /**
     * Create an invalid table range error.
     *
     * @param message   Error message describing the invalid range
     * @param sheetName Sheet name (if applicable)
     * @return BError
     */
    public static BError invalidTableRangeError(String message, String sheetName) {
        BMap<BString, Object> details = createErrorDetails(sheetName, null, null, null, null);
        return createTypedError(INVALID_TABLE_RANGE_ERROR_TYPE, message, details);
    }

    /**
     * Create an invalid table range error.
     *
     * @param message Error message describing the invalid range
     * @return BError
     */
    public static BError invalidTableRangeError(String message) {
        return createTypedError(INVALID_TABLE_RANGE_ERROR_TYPE, message, null);
    }

    private static BError createTypedError(String errorType, String message, BMap<BString, Object> details) {
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                errorType,
                StringUtils.fromString(message),
                null,
                details
        );
    }

    private static BMap<BString, Object> createErrorDetails(String sheetName, String tableName, String cellAddress,
                                                            Integer row, Integer column) {
        // Create typed ErrorDetails record instead of generic map
        BMap<BString, Object> details = ValueCreator.createRecordValue(ModuleUtils.getModule(), "ErrorDetails");

        if (sheetName != null) {
            details.put(StringUtils.fromString("sheetName"), StringUtils.fromString(sheetName));
        }
        if (tableName != null) {
            details.put(StringUtils.fromString("tableName"), StringUtils.fromString(tableName));
        }
        if (cellAddress != null) {
            details.put(StringUtils.fromString("cellAddress"), StringUtils.fromString(cellAddress));
        }
        if (row != null) {
            details.put(StringUtils.fromString("rowNumber"), (long) row);
        }
        if (column != null) {
            details.put(StringUtils.fromString("columnNumber"), (long) column);
        }

        return details;
    }

    private static String formatMessage(DiagnosticErrorCode code, Object... args) {
        String pattern = getErrorMessage(code);
        if (args.length > 0) {
            return MessageFormat.format(pattern, args);
        }
        return pattern;
    }

    private static String getErrorMessage(DiagnosticErrorCode code) {
        if (errorBundle != null) {
            try {
                return errorBundle.getString(ERROR_PREFIX + code.getMessageKey());
            } catch (Exception e) {
                // Fall through to default
            }
        }
        // Default message based on error code
        return code.getErrorCode() + ": " + code.getMessageKey().replace('.', ' ');
    }
}
