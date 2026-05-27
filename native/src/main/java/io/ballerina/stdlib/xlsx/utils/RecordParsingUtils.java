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

import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.creators.TypeCreator;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.flags.SymbolFlags;
import io.ballerina.runtime.api.types.ArrayType;
import io.ballerina.runtime.api.types.Field;
import io.ballerina.runtime.api.types.MapType;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.types.UnionType;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.stdlib.xlsx.xlsx.BallerinaErrorException;
import io.ballerina.stdlib.xlsx.xlsx.CellConverter;
import io.ballerina.stdlib.xlsx.xlsx.TypeConversionException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Utility class for record parsing operations shared between XlsxParser and SheetHandle.
 */
public final class RecordParsingUtils {

    private RecordParsingUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Data class for field mapping information.
     */
    public static class FieldMapping {
        public final String fieldName;
        public final Type type;
        public final boolean isOptional;

        public FieldMapping(String fieldName, Type type, boolean isOptional) {
            this.fieldName = fieldName;
            this.type = type;
            this.isOptional = isOptional;
        }
    }

    /**
     * Check if a field has the optional modifier (?).
     *
     * @param field The field to check
     * @return true if the field is optional
     */
    public static boolean isOptionalField(Field field) {
        return (field.getFlags() & SymbolFlags.OPTIONAL) == SymbolFlags.OPTIONAL;
    }

    /**
     * Check if a type is nilable (includes nil in union).
     *
     * @param type The type to check
     * @return true if the type is nilable
     */
    public static boolean isNilableType(Type type) {
        if (type.getTag() == TypeTags.UNION_TAG) {
            UnionType unionType = (UnionType) type;
            for (Type memberType : unionType.getMemberTypes()) {
                if (memberType.getTag() == TypeTags.NULL_TAG) {
                    return true;
                }
            }
        }
        return type.isNilable();
    }

    /**
     * Build a map of header names to column indices.
     *
     * @param headerRow       The row containing headers
     * @param usedRange       The used range of the sheet
     * @param caseInsensitive Whether to use case-insensitive matching
     * @return Map of header name to column index
     */
    public static Map<String, Integer> buildHeaderMap(Row headerRow, CellRangeAddress usedRange,
                                                       boolean caseInsensitive) {
        Map<String, Integer> headerMap = caseInsensitive ?
                new TreeMap<>(String.CASE_INSENSITIVE_ORDER) : new HashMap<>();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            Cell cell = headerRow.getCell(colIdx);
            if (cell != null) {
                String headerValue = cell.getStringCellValue();
                if (headerValue != null && !headerValue.trim().isEmpty()) {
                    String trimmed = headerValue.trim();
                    Integer prior = headerMap.put(trimmed, colIdx);
                    if (prior != null) {
                        // Two columns share the same header text. Fail loud rather than
                        // silently mapping the field to whichever column happens to come
                        // later in iteration order.
                        throw new BallerinaErrorException(DiagnosticLog.parseError(
                                "Duplicate header '" + trimmed + "' at columns "
                                        + prior + " and " + colIdx));
                    }
                }
            }
        }

        return headerMap;
    }

    /**
     * Build a header map from column indices for header-less sheets.
     * Generates column names as "col0", "col1", "col2", etc.
     *
     * @param usedRange       The used range of the sheet
     * @param caseInsensitive Whether to use case-insensitive matching
     * @return Map of generated header name to column index
     */
    public static Map<String, Integer> buildHeaderMapFromIndices(CellRangeAddress usedRange,
                                                                  boolean caseInsensitive) {
        Map<String, Integer> headerMap = caseInsensitive ?
                new TreeMap<>(String.CASE_INSENSITIVE_ORDER) : new HashMap<>();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            // Generate column name as "col0", "col1", etc. (relative to startCol)
            String colName = "col" + (colIdx - startCol);
            headerMap.put(colName, colIdx);
        }

        return headerMap;
    }

    /**
     * Convert column index and row index to Excel cell address (e.g., "A1", "B5").
     *
     * @param colIdx Column index (0-based)
     * @param rowIdx Row index (0-based)
     * @return Cell address string
     */
    public static String getCellAddress(int colIdx, int rowIdx) {
        StringBuilder colName = new StringBuilder();
        int col = colIdx;
        while (col >= 0) {
            colName.insert(0, (char) ('A' + (col % 26)));
            col = col / 26 - 1;
        }
        return colName.toString() + (rowIdx + 1);
    }

    /**
     * Build field-to-column mappings from a record type.
     *
     * @param recordType    The record type
     * @param headerMap     Map of header names to column indices
     * @param columnToField Output map of column index to field mapping
     * @param absentFields  Output list of fields without matching columns
     */
    public static void buildFieldMappings(
            RecordType recordType,
            Map<String, Integer> headerMap,
            Map<Integer, FieldMapping> columnToField,
            List<FieldMapping> absentFields) {

        Map<String, Field> fields = recordType.getFields();

        // Track which header each field claims, so two fields can't silently route
        // to the same column via duplicate @xlsx:Name annotations.
        Map<String, String> headerToField = new HashMap<>();

        for (Map.Entry<String, Field> entry : fields.entrySet()) {
            String fieldName = entry.getKey();
            Field field = entry.getValue();
            boolean isOptional = isOptionalField(field);

            // Check for @xlsx:Name annotation
            String headerName = AnnotationUtils.getHeaderName(recordType, fieldName);
            String priorField = headerToField.put(headerName, fieldName);
            if (priorField != null) {
                throw new BallerinaErrorException(DiagnosticLog.parseError(
                        "Fields '" + priorField + "' and '" + fieldName
                                + "' both resolve to header '" + headerName
                                + "' (check for duplicate @xlsx:Name annotations)"));
            }

            Integer colIndex = headerMap.get(headerName);
            if (colIndex != null) {
                columnToField.put(colIndex, new FieldMapping(fieldName, field.getFieldType(), isOptional));
            } else {
                // Field has no matching column in the sheet
                absentFields.add(new FieldMapping(fieldName, field.getFieldType(), isOptional));
            }
        }
    }

    /**
     * Check if a record type is open (has a rest field).
     * Open records can accept additional fields beyond those explicitly defined.
     *
     * @param recordType The record type to check
     * @return true if the record is open (has rest field), false if closed
     */
    public static boolean isOpenRecord(RecordType recordType) {
        // A record is open if it has a rest field type that is not null
        // Closed records (record {| ... |}) have no rest field
        Type restFieldType = recordType.getRestFieldType();
        return restFieldType != null && restFieldType.getTag() != TypeTags.NEVER_TAG;
    }

    /**
     * Build mappings for extra columns that don't match any defined field.
     * Used for populating open records with additional columns.
     *
     * @param headerMap      Map of header names to column indices
     * @param columnToField  Map of column indices already mapped to fields
     * @param extraColumns   Output map of column index to header name for unmapped columns
     */
    public static void buildExtraColumnMappings(
            Map<String, Integer> headerMap,
            Map<Integer, FieldMapping> columnToField,
            Map<Integer, String> extraColumns) {

        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            String headerName = entry.getKey();
            Integer colIndex = entry.getValue();

            // If this column is not mapped to any field, it's an extra column
            if (!columnToField.containsKey(colIndex)) {
                extraColumns.put(colIndex, headerName);
            }
        }
    }

    /**
     * Validate absent fields based on data projection settings.
     *
     * @param absentFields List of fields without matching columns
     * @param sheetName    Name of the sheet (for error messages)
     * @param config       Parsing configuration
     * @throws RuntimeException if validation fails
     */
    public static void validateAbsentFields(
            List<FieldMapping> absentFields,
            String sheetName,
            XlsxConfig config) {

        if (!config.isAllowDataProjection()) {
            // Strict mode: all fields must have matching columns
            if (!absentFields.isEmpty()) {
                StringBuilder missingFields = new StringBuilder();
                for (FieldMapping f : absentFields) {
                    if (missingFields.length() > 0) {
                        missingFields.append(", ");
                    }
                    missingFields.append(f.fieldName);
                }
                throw new BallerinaErrorException(DiagnosticLog.parseError(
                        "Data projection disabled but record has fields without matching columns: " +
                                missingFields.toString(),
                        sheetName, null, null));
            }
        } else if (!config.isAbsentAsNilableType()) {
            // Projection enabled but absentAsNilableType is false
            // Check if any absent field is required (not optional and not nilable)
            for (FieldMapping f : absentFields) {
                if (!f.isOptional && !isNilableType(f.type)) {
                    throw new BallerinaErrorException(DiagnosticLog.parseError(
                            "Required field '" + f.fieldName + "' has no matching column in the sheet. " +
                                    "Set 'absentAsNilableType: true' to allow absent columns for nilable types.",
                            sheetName, null, null));
                }
            }
        }
    }

    /**
     * Handle absent fields when building a record.
     * Sets absent fields to null if they're nilable or optional.
     *
     * @param record       The record being built
     * @param absentFields List of absent fields
     * @param config       Parsing configuration
     */
    public static void handleAbsentFields(
            BMap<BString, Object> record,
            List<FieldMapping> absentFields,
            XlsxConfig config) {

        if (config.isAbsentAsNilableType()) {
            for (FieldMapping absentField : absentFields) {
                // Set absent fields to null if they're nilable or optional
                if (isNilableType(absentField.type) || absentField.isOptional) {
                    record.put(StringUtils.fromString(absentField.fieldName), null);
                }
            }
        }
    }

    // =========================================================================
    // Shared Row Parsing Methods
    // =========================================================================

    /**
     * Context for row parsing operations.
     * Encapsulates all parameters needed for parsing rows from a sheet.
     */
    public static class ParseContext {
        public final Sheet sheet;
        public final CellRangeAddress usedRange;
        public final XlsxConfig config;
        public final RecordType recordType;
        public final MapType mapType;

        // Fail-safe support (optional - null if not enabled)
        public final Environment env;
        public final AtomicBoolean isOverwritten;

        /**
         * Constructor for record parsing.
         */
        public ParseContext(Sheet sheet, CellRangeAddress usedRange, XlsxConfig config,
                            RecordType recordType, Environment env, AtomicBoolean isOverwritten) {
            this.sheet = sheet;
            this.usedRange = usedRange;
            this.config = config;
            this.recordType = recordType;
            this.mapType = null;
            this.env = env;
            this.isOverwritten = isOverwritten;
        }

        /**
         * Constructor for map parsing.
         */
        public ParseContext(Sheet sheet, CellRangeAddress usedRange, XlsxConfig config,
                            MapType mapType, Environment env, AtomicBoolean isOverwritten) {
            this.sheet = sheet;
            this.usedRange = usedRange;
            this.config = config;
            this.recordType = null;
            this.mapType = mapType;
            this.env = env;
            this.isOverwritten = isOverwritten;
        }

        /**
         * Check if fail-safe mode is enabled.
         */
        public boolean isFailSafeEnabled() {
            return env != null && config.isFailSafeEnabled();
        }
    }

    /**
     * Parse sheet rows to record array.
     * Shared implementation used by both XlsxParser and SheetHandle.
     *
     * @param context Parse context containing sheet, config, and type info
     * @return BArray of records on success, or BError on failure
     */
    @SuppressWarnings("unchecked")
    public static Object parseRowsToRecords(ParseContext context) {
        ArrayType arrayType = TypeCreator.createArrayType(context.recordType);

        // Empty sheet check
        if (context.usedRange == null) {
            return ValueCreator.createArrayValue(arrayType);
        }

        // Build header-to-column mapping
        Map<String, Integer> headerMap;

        if (context.config.hasHeaders()) {
            // Header-based parsing: read headers from specified row
            Integer headerRowIndex = context.config.getHeaderRowIndex();
            Row headerRow = context.sheet.getRow(headerRowIndex);

            if (headerRow == null) {
                return DiagnosticLog.parseError(
                        "Header row " + headerRowIndex + " is empty",
                        context.sheet.getSheetName(), headerRowIndex, null);
            }

            headerMap = buildHeaderMap(headerRow, context.usedRange,
                    context.config.isCaseInsensitiveHeaders());
        } else {
            // Header-less parsing: generate column names as col0, col1, col2, ...
            headerMap = buildHeaderMapFromIndices(context.usedRange,
                    context.config.isCaseInsensitiveHeaders());
        }

        // Build field mappings
        Map<Integer, FieldMapping> columnToField = new HashMap<>();
        List<FieldMapping> absentFields = new ArrayList<>();
        buildFieldMappings(context.recordType, headerMap, columnToField, absentFields);

        // Build extra column mappings for open records
        Map<Integer, String> extraColumns = new HashMap<>();
        boolean isOpenRecord = isOpenRecord(context.recordType);
        Type restFieldType = null;
        if (isOpenRecord) {
            buildExtraColumnMappings(headerMap, columnToField, extraColumns);
            restFieldType = context.recordType.getRestFieldType();
        }

        // Validate absent fields
        try {
            validateAbsentFields(absentFields, context.sheet.getSheetName(), context.config);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }

        // Parse data rows
        int dataStartRow = context.config.getDataStartRowIndex();
        int endRow = context.usedRange.getLastRow();

        // Apply rowCount limit if set
        Integer rowCountLimit = context.config.getRowCount();

        List<BMap<BString, Object>> records = new ArrayList<>();
        int parsedCount = 0;

        for (int rowIdx = dataStartRow; rowIdx <= endRow; rowIdx++) {
            // Check row count limit (counts actual parsed rows, not all rows)
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = context.sheet.getRow(rowIdx);

            // Parse single row
            Object result = parseRowToRecord(row, rowIdx, columnToField, absentFields,
                    extraColumns, restFieldType, context);

            if (result instanceof BError) {
                return result;  // Fail-fast: return error
            }
            if (result == null) {
                continue;  // Fail-safe: row was skipped
            }
            records.add((BMap<BString, Object>) result);
            parsedCount++;
        }

        // Build result array
        BArray resultArray = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> record : records) {
            resultArray.append(record);
        }
        return resultArray;
    }

    /**
     * Parse a single row to a record.
     *
     * @param row            The Excel row to parse
     * @param rowIdx         The row index (for error reporting)
     * @param columnToField  Map of column index to field mapping
     * @param absentFields   List of fields without matching columns
     * @param extraColumns   Map of extra column indices to header names (for open records)
     * @param restFieldType  The rest field type for open records (null for closed records)
     * @param context        Parse context
     * @return BMap on success, BError on fail-fast error, or null on fail-safe skip
     */
    private static Object parseRowToRecord(Row row, int rowIdx,
            Map<Integer, FieldMapping> columnToField,
            List<FieldMapping> absentFields,
            Map<Integer, String> extraColumns,
            Type restFieldType,
            ParseContext context) {

        BMap<BString, Object> record = ValueCreator.createRecordValue(context.recordType);
        int currentColIdx = 0;

        try {
            // Parse defined fields
            for (Map.Entry<Integer, FieldMapping> entry : columnToField.entrySet()) {
                currentColIdx = entry.getKey();
                FieldMapping mapping = entry.getValue();

                Cell cell = row != null ? row.getCell(currentColIdx) : null;
                Object value = CellConverter.convert(cell, mapping.type, context.config);

                if (value != null) {
                    record.put(StringUtils.fromString(mapping.fieldName), value);
                } else {
                    // Nil handling
                    if (context.config.isNilAsOptionalField() && mapping.isOptional) {
                        continue;
                    }
                    if (isNilableType(mapping.type) || mapping.isOptional) {
                        record.put(StringUtils.fromString(mapping.fieldName), null);
                    } else {
                        // Blank cell for required non-nilable field is an error
                        throw new TypeConversionException(
                                "Required field '" + mapping.fieldName + "' cannot be null (blank cell)",
                                "", mapping.type.toString(), "null");
                    }
                }
            }

            // Parse extra columns for open records
            if (restFieldType != null && !extraColumns.isEmpty()) {
                for (Map.Entry<Integer, String> entry : extraColumns.entrySet()) {
                    currentColIdx = entry.getKey();
                    String headerName = entry.getValue();

                    Cell cell = row != null ? row.getCell(currentColIdx) : null;
                    Object value = CellConverter.convert(cell, restFieldType, context.config);

                    if (value != null) {
                        record.put(StringUtils.fromString(headerName), value);
                    } else if (!context.config.isNilAsOptionalField()) {
                        // Include nil values unless nilAsOptionalField is true
                        // Rest fields are always optional, so nil is acceptable
                        record.put(StringUtils.fromString(headerName), null);
                    }
                }
            }

            // Handle absent fields
            handleAbsentFields(record, absentFields, context.config);

            // Constraint validation (if enabled)
            BError validationError = performConstraintValidation(record, context.recordType,
                    context.config.isEnableConstraintValidation());
            if (validationError != null) {
                if (!context.isFailSafeEnabled()) {
                    return DiagnosticLog.constraintValidationError(validationError.getMessage(), rowIdx, null);
                }
                handleFailSafe(context, row, rowIdx, currentColIdx, validationError.getMessage());
                return null;
            }

            return record;

        } catch (TypeConversionException e) {
            String cellAddress = getCellAddress(currentColIdx, rowIdx);
            if (!context.isFailSafeEnabled()) {
                return DiagnosticLog.typeConversionError(e.getMessage(), cellAddress, rowIdx, currentColIdx);
            }
            // Fail-safe: log and skip
            handleFailSafe(context, row, rowIdx, currentColIdx, e.getMessage());
            return null;

        } catch (Exception e) {
            // TypeConversionException already handled above, this catches other exceptions
            if (!context.isFailSafeEnabled() || !FailSafeUtils.isAllowedFailSafe(e)) {
                throw e;
            }
            handleFailSafe(context, row, rowIdx, currentColIdx, e.getMessage());
            return null;
        }
    }

    /**
     * Perform constraint validation on a record.
     *
     * @param record    The record to validate
     * @param recordType The record type for validation
     * @param isEnabled Whether constraint validation is enabled
     * @return BError if validation fails, null if passes or disabled
     */
    private static BError performConstraintValidation(BMap<BString, Object> record,
                                                       RecordType recordType,
                                                       boolean isEnabled) {
        if (!isEnabled) {
            return null;
        }
        Object result = ConstraintUtils.validate(record, recordType);
        return (result instanceof BError) ? (BError) result : null;
    }

    /**
     * Parse sheet rows to map array.
     * Shared implementation used by both XlsxParser and SheetHandle.
     *
     * @param context Parse context containing sheet, config, and type info
     * @return BArray of maps on success, or BError on failure
     */
    @SuppressWarnings("unchecked")
    public static Object parseRowsToMaps(ParseContext context) {
        ArrayType arrayType = TypeCreator.createArrayType(context.mapType);

        // Empty sheet check
        if (context.usedRange == null) {
            return ValueCreator.createArrayValue(arrayType);
        }

        // Build header map
        Map<String, Integer> headerMap;

        if (context.config.hasHeaders()) {
            // Header-based parsing: read headers from specified row
            Integer headerRowIndex = context.config.getHeaderRowIndex();
            Row headerRow = context.sheet.getRow(headerRowIndex);

            if (headerRow == null) {
                return DiagnosticLog.parseError(
                        "Header row " + headerRowIndex + " is empty",
                        context.sheet.getSheetName(), headerRowIndex, null);
            }

            headerMap = buildHeaderMap(headerRow, context.usedRange,
                    context.config.isCaseInsensitiveHeaders());
        } else {
            // Header-less parsing: generate column names as col0, col1, col2, ...
            headerMap = buildHeaderMapFromIndices(context.usedRange,
                    context.config.isCaseInsensitiveHeaders());
        }
        Map<Integer, String> columnToHeader = new HashMap<>();
        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            columnToHeader.put(entry.getValue(), entry.getKey());
        }

        // Parse data rows
        int dataStartRow = context.config.getDataStartRowIndex();
        int endRow = context.usedRange.getLastRow();
        Type constraintType = context.mapType.getConstrainedType();

        // Apply rowCount limit if set
        Integer rowCountLimit = context.config.getRowCount();

        List<BMap<BString, Object>> maps = new ArrayList<>();
        int parsedCount = 0;

        for (int rowIdx = dataStartRow; rowIdx <= endRow; rowIdx++) {
            // Check row count limit (counts actual parsed rows, not all rows)
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = context.sheet.getRow(rowIdx);

            // Parse single row
            Object result = parseRowToMap(row, rowIdx, columnToHeader, constraintType, context);

            if (result instanceof BError) {
                return result;  // Fail-fast: return error
            }
            if (result == null) {
                continue;  // Fail-safe: row was skipped
            }
            maps.add((BMap<BString, Object>) result);
            parsedCount++;
        }

        // Build result array
        BArray resultArray = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> map : maps) {
            resultArray.append(map);
        }
        return resultArray;
    }

    /**
     * Parse a single row to a map.
     *
     * @return BMap on success, BError on fail-fast error, or null on fail-safe skip
     */
    private static Object parseRowToMap(Row row, int rowIdx,
            Map<Integer, String> columnToHeader, Type constraintType, ParseContext context) {

        BMap<BString, Object> map = ValueCreator.createMapValue(context.mapType);
        int currentColIdx = 0;

        try {
            for (Map.Entry<Integer, String> entry : columnToHeader.entrySet()) {
                currentColIdx = entry.getKey();
                String header = entry.getValue();

                Cell cell = row != null ? row.getCell(currentColIdx) : null;
                Object value = CellConverter.convert(cell, constraintType, context.config);

                if (value != null) {
                    map.put(StringUtils.fromString(header), value);
                } else if (!context.config.isNilAsOptionalField()) {
                    // Include nil values unless nilAsOptionalField is true
                    map.put(StringUtils.fromString(header), null);
                }
            }

            return map;

        } catch (TypeConversionException e) {
            String cellAddress = getCellAddress(currentColIdx, rowIdx);
            if (!context.isFailSafeEnabled()) {
                return DiagnosticLog.typeConversionError(e.getMessage(), cellAddress, rowIdx, currentColIdx);
            }
            // Fail-safe: log and skip
            handleFailSafe(context, row, rowIdx, currentColIdx, e.getMessage());
            return null;

        } catch (Exception e) {
            // TypeConversionException already handled above, this catches other exceptions
            if (!context.isFailSafeEnabled() || !FailSafeUtils.isAllowedFailSafe(e)) {
                throw e;
            }
            handleFailSafe(context, row, rowIdx, currentColIdx, e.getMessage());
            return null;
        }
    }

    /**
     * Handle fail-safe logging for a row parsing error.
     */
    private static void handleFailSafe(ParseContext context, Row row,
            int rowIdx, int colIdx, String message) {
        if (context.env == null) {
            return;  // Fail-safe not enabled
        }

        String offendingRow = rowToJsonArray(row, context.usedRange);
        FailSafeUtils.handleFailSafeLogging(
                context.env, context.config.getFailSafe(),
                new RuntimeException(message), offendingRow,
                rowIdx, colIdx, context.isOverwritten,
                context.config.isEnableConsoleLogs(), context.config.isIncludeSourceDataInConsole()
        );
    }

    /**
     * Convert a row to a JSON array string for fail-safe logging.
     * Example output: ["John", "abc", "Sales", "30"]
     *
     * @param row       The Excel row
     * @param usedRange The used range to determine column bounds
     * @return JSON array string representation of the row
     */
    public static String rowToJsonArray(Row row, CellRangeAddress usedRange) {
        if (row == null) {
            return "[]";
        }

        StringBuilder json = new StringBuilder("[");
        boolean first = true;
        int startCol = usedRange != null ? usedRange.getFirstColumn() : 0;
        int endCol = usedRange != null ? usedRange.getLastColumn() : row.getLastCellNum() - 1;

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            if (!first) {
                json.append(", ");
            }
            first = false;

            Cell cell = row.getCell(colIdx);
            String value = "";
            if (cell != null) {
                try {
                    value = CellConverter.convertToStringRaw(cell);
                } catch (Exception e) {
                    value = "[error]";
                }
            }
            // Escape special characters (quotes, newlines, tabs, etc.) and wrap in quotes
            json.append("\"").append(FailSafeUtils.escapeJson(value)).append("\"");
        }
        json.append("]");

        return json.toString();
    }
}
