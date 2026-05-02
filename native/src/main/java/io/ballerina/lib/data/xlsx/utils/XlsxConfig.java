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

package io.ballerina.lib.data.xlsx.utils;

import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

/**
 * Configuration holder for XLSX parsing and writing options.
 */
public class XlsxConfig {

    // Parse options
    private Integer headerRowIndex = Constants.DEFAULT_HEADER_ROW;  // null = no headers (use col0, col1, ...)
    private Integer dataStartRowIndex;
    private Integer rowCount;  // null = read all rows
    private String formulaMode = Constants.FORMULA_MODE_CACHED;
    private boolean enableConstraintValidation = true;
    private boolean caseInsensitiveHeaders = false;

    // Data projection options
    private boolean allowDataProjection = true;  // false disables projection entirely
    private boolean nilAsOptionalField = false;
    private boolean absentAsNilableType = false;

    // Fail-safe options
    private BMap<?, ?> failSafe = null;
    private boolean enableConsoleLogs = true;
    private boolean includeSourceDataInConsole = false;

    // Write options
    private String writeSheetName = Constants.DEFAULT_SHEET_NAME;
    private boolean writeHeaders = true;
    private int startRowIndex = 0;

    /**
     * Create config from Ballerina parse options map.
     *
     * @param options Ballerina options map
     * @return XlsxConfig instance
     */
    public static XlsxConfig fromParseOptions(BMap<BString, Object> options) {
        XlsxConfig config = new XlsxConfig();

        if (options == null) {
            return config;
        }

        // Header configuration - headerRowIndex can be null (no headers)
        Object headerRowIndexVal = options.get(Constants.HEADER_ROW_INDEX);
        if (headerRowIndexVal == null) {
            config.headerRowIndex = null;  // No headers - use col0, col1, ...
        } else {
            config.headerRowIndex = ((Long) headerRowIndexVal).intValue();
        }

        Object dataStartRowIndexVal = options.get(Constants.DATA_START_ROW_INDEX);
        if (dataStartRowIndexVal != null) {
            config.dataStartRowIndex = ((Long) dataStartRowIndexVal).intValue();
        }

        // Row count limit
        Object rowCountVal = options.get(Constants.ROW_COUNT);
        if (rowCountVal != null) {
            config.rowCount = ((Long) rowCountVal).intValue();
        }

        // Formula handling
        Object formulaModeVal = options.get(Constants.FORMULA_MODE);
        if (formulaModeVal != null) {
            config.formulaMode = formulaModeVal.toString();
        }

        // Constraint validation
        Object constraintVal = options.get(Constants.ENABLE_CONSTRAINT_VALIDATION);
        if (constraintVal != null) {
            config.enableConstraintValidation = (Boolean) constraintVal;
        }

        // Case-insensitive headers
        Object caseInsensitiveVal = options.get(Constants.CASE_INSENSITIVE_HEADERS);
        if (caseInsensitiveVal != null) {
            config.caseInsensitiveHeaders = (Boolean) caseInsensitiveVal;
        }

        // Data projection options
        Object projectionVal = options.get(Constants.ALLOW_DATA_PROJECTION);
        if (projectionVal != null) {
            if (projectionVal instanceof Boolean) {
                // allowDataProjection = false disables projection
                config.allowDataProjection = (Boolean) projectionVal;
            } else if (projectionVal instanceof BMap<?, ?>) {
                // allowDataProjection is a record with options
                config.allowDataProjection = true;
                BMap<?, ?> projectionOpts = (BMap<?, ?>) projectionVal;

                Object nilAsOptionalVal = projectionOpts.get(Constants.NIL_AS_OPTIONAL_FIELD);
                if (nilAsOptionalVal != null) {
                    config.nilAsOptionalField = (Boolean) nilAsOptionalVal;
                }

                Object absentAsNilableVal = projectionOpts.get(Constants.ABSENT_AS_NILABLE_TYPE);
                if (absentAsNilableVal != null) {
                    config.absentAsNilableType = (Boolean) absentAsNilableVal;
                }
            }
        }

        // Fail-safe options
        Object failSafeVal = options.get(Constants.FAIL_SAFE);
        if (failSafeVal instanceof BMap<?, ?>) {
            config.failSafe = (BMap<?, ?>) failSafeVal;

            // Extract console logging options
            Object enableConsoleLogsVal = config.failSafe.get(Constants.ENABLE_CONSOLE_LOGS);
            if (enableConsoleLogsVal != null) {
                config.enableConsoleLogs = (Boolean) enableConsoleLogsVal;
            }

            Object includeSourceDataVal = config.failSafe.get(Constants.INCLUDE_SOURCE_DATA_IN_CONSOLE);
            if (includeSourceDataVal != null) {
                config.includeSourceDataInConsole = (Boolean) includeSourceDataVal;
            }
        }

        return config;
    }

    /**
     * Create config from Ballerina write options map.
     *
     * @param options Ballerina options map
     * @return XlsxConfig instance
     */
    public static XlsxConfig fromWriteOptions(BMap<BString, Object> options) {
        XlsxConfig config = new XlsxConfig();

        if (options == null) {
            return config;
        }

        Object sheetNameVal = options.get(Constants.WRITE_SHEET_NAME);
        if (sheetNameVal != null) {
            config.writeSheetName = sheetNameVal.toString();
        }

        Object writeHeadersVal = options.get(Constants.WRITE_HEADERS);
        if (writeHeadersVal != null) {
            config.writeHeaders = (Boolean) writeHeadersVal;
        }

        Object startRowIndexVal = options.get(Constants.START_ROW_INDEX);
        if (startRowIndexVal != null) {
            config.startRowIndex = ((Long) startRowIndexVal).intValue();
        }

        return config;
    }

    // Getters

    /**
     * Get the header row index, or null if no headers (header-less mode).
     */
    public Integer getHeaderRowIndex() {
        return headerRowIndex;
    }

    /**
     * Check if the sheet has headers.
     */
    public boolean hasHeaders() {
        return headerRowIndex != null;
    }

    /**
     * Get the data start row index.
     * Default: row after header, or row 0 if no headers.
     */
    public int getDataStartRowIndex() {
        if (dataStartRowIndex != null) {
            return dataStartRowIndex;
        }
        // Default: row after header, or 0 if no headers
        return headerRowIndex != null ? headerRowIndex + 1 : 0;
    }

    public boolean hasExplicitDataStartRowIndex() {
        return dataStartRowIndex != null;
    }

    /**
     * Get the row count limit, or null if no limit (read all rows).
     */
    public Integer getRowCount() {
        return rowCount;
    }

    /**
     * Check if a row count limit is set.
     */
    public boolean hasRowCountLimit() {
        return rowCount != null;
    }

    public String getFormulaMode() {
        return formulaMode;
    }

    public boolean isFormulaModeText() {
        return Constants.FORMULA_MODE_TEXT.equals(formulaMode);
    }

    public boolean isEnableConstraintValidation() {
        return enableConstraintValidation;
    }

    public boolean isCaseInsensitiveHeaders() {
        return caseInsensitiveHeaders;
    }

    // Data projection getters

    public boolean isAllowDataProjection() {
        return allowDataProjection;
    }

    public boolean isNilAsOptionalField() {
        return nilAsOptionalField;
    }

    public boolean isAbsentAsNilableType() {
        return absentAsNilableType;
    }

    public String getWriteSheetName() {
        return writeSheetName;
    }

    public boolean isWriteHeaders() {
        return writeHeaders;
    }

    public int getStartRowIndex() {
        return startRowIndex;
    }

    // Fail-safe getters

    public BMap<?, ?> getFailSafe() {
        return failSafe;
    }

    public boolean isFailSafeEnabled() {
        return failSafe != null;
    }

    public boolean isEnableConsoleLogs() {
        return enableConsoleLogs;
    }

    public boolean isIncludeSourceDataInConsole() {
        return includeSourceDataInConsole;
    }
}
