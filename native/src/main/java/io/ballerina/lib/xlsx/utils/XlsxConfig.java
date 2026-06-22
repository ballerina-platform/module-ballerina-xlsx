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

import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

/**
 * Configuration holder for XLSX parsing and writing options.
 */
public class XlsxConfig {

    // Parse options field names
    private static final BString HEADER_ROW_INDEX = StringUtils.fromString("headerRowIndex");
    private static final BString DATA_START_ROW_INDEX = StringUtils.fromString("dataStartRowIndex");
    private static final BString ROW_COUNT = StringUtils.fromString("rowCount");
    private static final BString FORMULA_MODE = StringUtils.fromString("formulaMode");
    private static final BString ENABLE_CONSTRAINT_VALIDATION = StringUtils.fromString("enableConstraintValidation");
    private static final BString CASE_INSENSITIVE_HEADERS = StringUtils.fromString("caseInsensitiveHeaders");

    // Data projection options field names
    private static final BString ALLOW_DATA_PROJECTION = StringUtils.fromString("allowDataProjection");
    private static final BString NIL_AS_OPTIONAL_FIELD = StringUtils.fromString("nilAsOptionalField");
    private static final BString ABSENT_AS_NILABLE_TYPE = StringUtils.fromString("absentAsNilableType");

    // Fail-safe options field names read at config time
    private static final BString FAIL_SAFE = StringUtils.fromString("failSafe");
    private static final BString ENABLE_CONSOLE_LOGS = StringUtils.fromString("enableConsoleLogs");
    private static final BString INCLUDE_SOURCE_DATA_IN_CONSOLE = StringUtils.fromString("includeSourceDataInConsole");

    // Write options field names. WRITE_SHEET_NAME is public so the sheet name positional
    // argument can be injected into the options map before extraction.
    public static final BString WRITE_SHEET_NAME = StringUtils.fromString("sheetName");
    private static final BString WRITE_HEADERS = StringUtils.fromString("writeHeaders");
    private static final BString START_ROW_INDEX = StringUtils.fromString("startRowIndex");
    private static final BString START_COLUMN_INDEX = StringUtils.fromString("startColumnIndex");
    public static final BString WRITE_SHEET_MODE = StringUtils.fromString("sheetWriteMode");
    private static final BString WRITE_TABLE_MODE = StringUtils.fromString("tableWriteMode");
    private static final BString TABLE_INSERT_AT = StringUtils.fromString("insertAt");

    // Formula mode values
    private static final String FORMULA_MODE_CACHED = "CACHED";
    private static final String FORMULA_MODE_TEXT = "TEXT";

    // Default values
    private static final String DEFAULT_SHEET_NAME = "Sheet1";
    private static final int DEFAULT_HEADER_ROW = 0;

    // Parse options
    private Integer headerRowIndex = DEFAULT_HEADER_ROW;  // null = no headers (use col0, col1, ...)
    private Integer dataStartRowIndex;
    private Integer rowCount;  // null = read all rows
    private String formulaMode = FORMULA_MODE_CACHED;
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
    private String writeSheetName = DEFAULT_SHEET_NAME;
    private boolean writeHeaders = true;
    private int startRowIndex = 0;
    // Boxed: null = no explicit startRowIndex (Sheet.putRows resolves the mode's natural point).
    // Carries the optional `int startRowIndex?` from WriteOptions without disturbing the primitive
    // startRowIndex used by writeSheet.
    private Integer startRowIndexOverride = null;
    private int startColumnIndex = 0;
    // A Ballerina SheetWriteMode enum member; default matches SheetWriteOptions.
    private String sheetWriteMode = "FAIL_IF_EXISTS";
    // A Ballerina TableWriteMode enum member; default matches TableWriteOptions.
    private String tableWriteMode = "REPLACE";
    // Boxed: null = no explicit insert position (APPEND appends at the bottom of the data).
    private Integer tableInsertAt = null;

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
        Object headerRowIndexVal = options.get(HEADER_ROW_INDEX);
        if (headerRowIndexVal == null) {
            config.headerRowIndex = null;  // No headers - use col0, col1, ...
        } else {
            config.headerRowIndex = ((Long) headerRowIndexVal).intValue();
        }

        Object dataStartRowIndexVal = options.get(DATA_START_ROW_INDEX);
        if (dataStartRowIndexVal != null) {
            config.dataStartRowIndex = ((Long) dataStartRowIndexVal).intValue();
        }

        // Row count limit
        Object rowCountVal = options.get(ROW_COUNT);
        if (rowCountVal != null) {
            config.rowCount = ((Long) rowCountVal).intValue();
        }

        // Formula handling
        Object formulaModeVal = options.get(FORMULA_MODE);
        if (formulaModeVal != null) {
            config.formulaMode = formulaModeVal.toString();
        }

        // Constraint validation
        Object constraintVal = options.get(ENABLE_CONSTRAINT_VALIDATION);
        if (constraintVal != null) {
            config.enableConstraintValidation = (Boolean) constraintVal;
        }

        // Case-insensitive headers
        Object caseInsensitiveVal = options.get(CASE_INSENSITIVE_HEADERS);
        if (caseInsensitiveVal != null) {
            config.caseInsensitiveHeaders = (Boolean) caseInsensitiveVal;
        }

        // Data projection options
        Object projectionVal = options.get(ALLOW_DATA_PROJECTION);
        if (projectionVal != null) {
            if (projectionVal instanceof Boolean) {
                // allowDataProjection = false disables projection
                config.allowDataProjection = (Boolean) projectionVal;
            } else if (projectionVal instanceof BMap<?, ?>) {
                // allowDataProjection is a record with options
                config.allowDataProjection = true;
                BMap<?, ?> projectionOpts = (BMap<?, ?>) projectionVal;

                Object nilAsOptionalVal = projectionOpts.get(NIL_AS_OPTIONAL_FIELD);
                if (nilAsOptionalVal != null) {
                    config.nilAsOptionalField = (Boolean) nilAsOptionalVal;
                }

                Object absentAsNilableVal = projectionOpts.get(ABSENT_AS_NILABLE_TYPE);
                if (absentAsNilableVal != null) {
                    config.absentAsNilableType = (Boolean) absentAsNilableVal;
                }
            }
        }

        // Fail-safe options
        Object failSafeVal = options.get(FAIL_SAFE);
        if (failSafeVal instanceof BMap<?, ?>) {
            config.failSafe = (BMap<?, ?>) failSafeVal;

            // Extract console logging options
            Object enableConsoleLogsVal = config.failSafe.get(ENABLE_CONSOLE_LOGS);
            if (enableConsoleLogsVal != null) {
                config.enableConsoleLogs = (Boolean) enableConsoleLogsVal;
            }

            Object includeSourceDataVal = config.failSafe.get(INCLUDE_SOURCE_DATA_IN_CONSOLE);
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

        Object sheetNameVal = options.get(WRITE_SHEET_NAME);
        if (sheetNameVal != null) {
            config.writeSheetName = sheetNameVal.toString();
        }

        Object writeHeadersVal = options.get(WRITE_HEADERS);
        if (writeHeadersVal != null) {
            config.writeHeaders = (Boolean) writeHeadersVal;
        }

        Object startRowIndexVal = options.get(START_ROW_INDEX);
        if (startRowIndexVal != null) {
            config.startRowIndex = ((Long) startRowIndexVal).intValue();
            config.startRowIndexOverride = ((Long) startRowIndexVal).intValue();
        }

        Object startColumnIndexVal = options.get(START_COLUMN_INDEX);
        if (startColumnIndexVal != null) {
            config.startColumnIndex = ((Long) startColumnIndexVal).intValue();
        }

        Object sheetWriteModeVal = options.get(WRITE_SHEET_MODE);
        if (sheetWriteModeVal != null) {
            config.sheetWriteMode = sheetWriteModeVal.toString();
        }

        Object tableWriteModeVal = options.get(WRITE_TABLE_MODE);
        if (tableWriteModeVal != null) {
            config.tableWriteMode = tableWriteModeVal.toString();
        }

        Object tableInsertAtVal = options.get(TABLE_INSERT_AT);
        if (tableInsertAtVal != null) {
            config.tableInsertAt = ((Long) tableInsertAtVal).intValue();
        }

        // headerRowIndex is carried by single-row write options (setRow) to locate the
        // header row for record/map column alignment.
        Object headerRowIndexVal = options.get(HEADER_ROW_INDEX);
        if (headerRowIndexVal != null) {
            config.headerRowIndex = ((Long) headerRowIndexVal).intValue();
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

    public boolean isFormulaModeText() {
        return FORMULA_MODE_TEXT.equals(formulaMode);
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

    /**
     * Whether Sheet.putRows was given an explicit startRowIndex (vs. leaving it `()` to resolve
     * the mode's natural point).
     */
    public boolean hasStartRowIndex() {
        return startRowIndexOverride != null;
    }

    public Integer getStartRowIndexOverride() {
        return startRowIndexOverride;
    }

    public int getStartColumnIndex() {
        return startColumnIndex;
    }

    public String getSheetWriteMode() {
        return sheetWriteMode;
    }

    public String getTableWriteMode() {
        return tableWriteMode;
    }

    /**
     * Explicit APPEND insert position (0-based data-row index), or null to append at the bottom.
     */
    public Integer getTableInsertAt() {
        return tableInsertAt;
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
