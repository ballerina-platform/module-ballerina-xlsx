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

package io.ballerina.lib.data.xlsx.xlsx;

import io.ballerina.lib.data.xlsx.utils.Constants;
import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.utils.ModuleUtils;
import io.ballerina.lib.data.xlsx.utils.RecordParsingUtils;
import io.ballerina.lib.data.xlsx.utils.RecordParsingUtils.FieldMapping;
import io.ballerina.lib.data.xlsx.utils.RowTypeUtils;
import io.ballerina.lib.data.xlsx.utils.UsedRangeDetector;
import io.ballerina.lib.data.xlsx.utils.XlsxConfig;
import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.creators.TypeCreator;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.ArrayType;
import io.ballerina.runtime.api.types.MapType;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BObject;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Native handle for Apache POI Sheet.
 * This class wraps the POI Sheet and provides methods called from Ballerina.
 */
public final class SheetHandle {

    private static final String SHEET_NATIVE_KEY = "sheetNative";

    private SheetHandle() {
        // Private constructor to prevent instantiation
    }

    /**
     * Initialize a Ballerina Sheet object with a POI Sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @param sheet    POI Sheet
     */
    static void initSheet(BObject sheetObj, Sheet sheet) {
        sheetObj.addNativeData(SHEET_NATIVE_KEY, sheet);
    }

    /**
     * Get the sheet name.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Sheet name
     */
    public static BString getName(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        return StringUtils.fromString(sheet.getSheetName());
    }

    /**
     * Get the used range of the sheet in A1 notation.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Used range string (e.g., "A1:D50")
     */
    public static BString getUsedRange(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return StringUtils.fromString(UsedRangeDetector.toA1Notation(range));
    }

    /**
     * Get the used cell range of the sheet as a structured record.
     *
     * @param sheetObj Ballerina Sheet object
     * @return CellRange record with 0-based indices, or null if sheet is empty
     */
    public static Object getUsedCellRange(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);

        if (range == null) {
            return null;
        }

        // Create CellRange record using the module's record type
        BMap<BString, Object> cellRange = ValueCreator.createRecordValue(
                ModuleUtils.getModule(), "CellRange");
        cellRange.put(StringUtils.fromString("firstRowIndex"), (long) range.getFirstRow());
        cellRange.put(StringUtils.fromString("lastRowIndex"), (long) range.getLastRow());
        cellRange.put(StringUtils.fromString("firstColumnIndex"), (long) range.getFirstColumn());
        cellRange.put(StringUtils.fromString("lastColumnIndex"), (long) range.getLastColumn());

        return cellRange;
    }

    /**
     * Get the row count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of rows with data
     */
    public static long getRowCount(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return UsedRangeDetector.getRowCount(range);
    }

    /**
     * Get the column count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of columns with data
     */
    public static long getColumnCount(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return UsedRangeDetector.getColumnCount(range);
    }

    /**
     * Get rows from the sheet.
     *
     * @param env        Ballerina runtime environment (for fail-safe support)
     * @param sheetObj   Ballerina Sheet object
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Array of rows (string[][] or record[])
     */
    public static Object getRows(Environment env, BObject sheetObj, BMap<BString, Object> options,
                                  BTypedesc targetType) {
        Sheet sheet = getSheet(sheetObj);
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try {
            Type describingType = targetType.getDescribingType();
            int typeTag = describingType.getTag();

            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) describingType;
                Type elementType = arrayType.getElementType();
                // Resolve referenced types (important for module-defined types)
                Type resolvedElementType = TypeUtils.getReferredType(elementType);
                int elementTag = resolvedElementType.getTag();

                // string[][]
                if (elementTag == TypeTags.ARRAY_TAG) {
                    return getRowsAsStringArray(sheet, config);
                }

                // record[]
                if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                    RecordType recordType = (RecordType) resolvedElementType;

                    // Check if this is a Row-wrapped type (has rowIndex and value fields)
                    if (RowTypeUtils.isRowWrappedType(recordType)) {
                        return getRowsAsRowWrappedRecords(env, sheet, config, recordType);
                    }

                    return getRowsAsRecords(env, sheet, config, recordType);
                }

                // map<anydata>[]
                if (elementTag == TypeTags.MAP_TAG) {
                    return getRowsAsMaps(env, sheet, config, (MapType) resolvedElementType);
                }
            }

            // Default: string[][]
            return getRowsAsStringArray(sheet, config);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting rows: " + e.getMessage(), e);
        }
    }

    /**
     * Get a single row from the sheet by index.
     *
     * @param env        Ballerina runtime environment
     * @param sheetObj   Ballerina Sheet object
     * @param index      Row index (0-based, relative to data start row)
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Single row (string[] or record{})
     */
    public static Object getRow(Environment env, BObject sheetObj, long index, BMap<BString, Object> options,
                                 BTypedesc targetType) {
        Sheet sheet = getSheet(sheetObj);
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try {
            CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

            if (usedRange == null) {
                return DiagnosticLog.error("Sheet is empty, cannot get row at index " + index);
            }

            int dataStartRow = config.getDataStartRowIndex();
            int actualRowIndex = dataStartRow + (int) index;
            int endRow = usedRange.getLastRow();

            if (actualRowIndex < dataStartRow || actualRowIndex > endRow) {
                return DiagnosticLog.error("Row index " + index + " out of range (0-" +
                        (endRow - dataStartRow) + ")");
            }

            Row row = sheet.getRow(actualRowIndex);

            Type describingType = targetType.getDescribingType();
            // Resolve referenced types (important for module-defined types)
            Type resolvedType = TypeUtils.getReferredType(describingType);
            int typeTag = resolvedType.getTag();

            // string[] - single row as string array
            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) resolvedType;
                Type elementType = TypeUtils.getReferredType(arrayType.getElementType());
                if (elementType.getTag() == TypeTags.STRING_TAG) {
                    return getRowAsStringArray(row, usedRange, config);
                }
            }

            // record{} - single row as record
            if (typeTag == TypeTags.RECORD_TYPE_TAG) {
                return getRowAsRecord(sheet, row, usedRange, config, (RecordType) resolvedType);
            }

            // Default: string[]
            return getRowAsStringArray(row, usedRange, config);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting row: " + e.getMessage(), e);
        }
    }

    /**
     * Get a single row as string[].
     */
    private static BArray getRowAsStringArray(Row row, CellRangeAddress usedRange, XlsxConfig config) {
        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        BArray rowArray = ValueCreator.createArrayValue(stringType);

        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            Cell cell = row != null ? row.getCell(colIdx) : null;
            String value = CellConverter.convertToString(cell, config);
            rowArray.append(StringUtils.fromString(value));
        }

        return rowArray;
    }

    /**
     * Get a single row as record.
     */
    private static BMap<BString, Object> getRowAsRecord(Sheet sheet, Row row, CellRangeAddress usedRange,
                                                         XlsxConfig config, RecordType recordType) {
        // Build header-to-column mapping
        Map<String, Integer> headerMap;

        if (config.hasHeaders()) {
            // Header-based parsing: read headers from specified row
            Integer headerRowIndex = config.getHeaderRowIndex();
            Row headerRow = sheet.getRow(headerRowIndex);

            if (headerRow == null) {
                throw new BallerinaErrorException(
                        DiagnosticLog.parseError("Header row " + headerRowIndex + " is empty"));
            }

            headerMap = RecordParsingUtils.buildHeaderMap(headerRow, usedRange, config.isCaseInsensitiveHeaders());
        } else {
            // Header-less parsing: generate column names as col0, col1, col2, ...
            headerMap = RecordParsingUtils.buildHeaderMapFromIndices(usedRange, config.isCaseInsensitiveHeaders());
        }

        // Get field mappings and track fields without matching columns
        Map<Integer, FieldMapping> columnToField = new HashMap<>();
        List<FieldMapping> absentFields = new ArrayList<>();
        RecordParsingUtils.buildFieldMappings(recordType, headerMap, columnToField, absentFields);

        // Build extra column mappings for open records
        Map<Integer, String> extraColumns = new HashMap<>();
        boolean isOpenRecord = RecordParsingUtils.isOpenRecord(recordType);
        Type restFieldType = null;
        if (isOpenRecord) {
            RecordParsingUtils.buildExtraColumnMappings(headerMap, columnToField, extraColumns);
            restFieldType = recordType.getRestFieldType();
        }

        // Validate absent fields based on data projection settings
        RecordParsingUtils.validateAbsentFields(absentFields, sheet.getSheetName(), config);

        // Create record from row
        BMap<BString, Object> record = ValueCreator.createRecordValue(recordType);

        // Parse defined fields
        for (Map.Entry<Integer, FieldMapping> entry : columnToField.entrySet()) {
            int colIdx = entry.getKey();
            FieldMapping mapping = entry.getValue();

            Cell cell = row != null ? row.getCell(colIdx) : null;
            Object value;
            try {
                value = CellConverter.convert(cell, mapping.type, config);
            } catch (TypeConversionException e) {
                int rowIdx = row != null ? row.getRowNum() : -1;
                String cellAddress = RecordParsingUtils.getCellAddress(colIdx, rowIdx);
                throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                        e.getMessage(), cellAddress, rowIdx, colIdx));
            }

            if (value != null) {
                record.put(StringUtils.fromString(mapping.fieldName), value);
            } else {
                // Value is nil - handle based on projection settings
                if (config.isNilAsOptionalField() && mapping.isOptional) {
                    continue;
                }
                if (RecordParsingUtils.isNilableType(mapping.type) || mapping.isOptional) {
                    record.put(StringUtils.fromString(mapping.fieldName), null);
                } else {
                    // Blank cell for required non-nilable field is an error
                    int rowIdx = row != null ? row.getRowNum() : -1;
                    String cellAddress = RecordParsingUtils.getCellAddress(colIdx, rowIdx);
                    throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                            "Required field '" + mapping.fieldName + "' cannot be null (blank cell)",
                            cellAddress, rowIdx, colIdx));
                }
            }
        }

        // Parse extra columns for open records
        if (restFieldType != null && !extraColumns.isEmpty()) {
            for (Map.Entry<Integer, String> entry : extraColumns.entrySet()) {
                int colIdx = entry.getKey();
                String headerName = entry.getValue();

                Cell cell = row != null ? row.getCell(colIdx) : null;
                Object value;
                try {
                    value = CellConverter.convert(cell, restFieldType, config);
                } catch (TypeConversionException e) {
                    int rowIdx = row != null ? row.getRowNum() : -1;
                    String cellAddress = RecordParsingUtils.getCellAddress(colIdx, rowIdx);
                    throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                            e.getMessage(), cellAddress, rowIdx, colIdx));
                }

                if (value != null) {
                    record.put(StringUtils.fromString(headerName), value);
                } else if (!config.isNilAsOptionalField()) {
                    // Include nil values unless nilAsOptionalField is true
                    record.put(StringUtils.fromString(headerName), null);
                }
            }
        }

        // Handle absent fields
        RecordParsingUtils.handleAbsentFields(record, absentFields, config);

        return record;
    }

    /**
     * Put rows into the sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @param data     Data to write
     * @param options  Write options
     * @return null on success, error on failure
     */
    public static Object putRows(BObject sheetObj, BArray data, BMap<BString, Object> options) {
        Sheet sheet = getSheet(sheetObj);
        return XlsxWriter.writeToSheet(sheet, data, options);
    }

    /**
     * Get rows as string[][].
     */
    private static BArray getRowsAsStringArray(Sheet sheet, XlsxConfig config) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        ArrayType resultType = TypeCreator.createArrayType(stringType);

        if (usedRange == null) {
            return ValueCreator.createArrayValue(resultType);
        }

        int startRow = config.hasExplicitDataStartRowIndex() ?
                config.getDataStartRowIndex() : usedRange.getFirstRow();
        int endRow = usedRange.getLastRow();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        // Apply rowCount limit if set
        Integer rowCountLimit = config.getRowCount();

        List<BArray> rows = new ArrayList<>();
        int parsedCount = 0;

        for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
            // Check row count limit (counts actual parsed rows, not all rows)
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = sheet.getRow(rowIdx);

            // Skip empty rows (Row wrapper type will override this behavior)
            if (UsedRangeDetector.isRowEmpty(row)) {
                continue;
            }

            BArray rowArray = ValueCreator.createArrayValue(stringType);
            for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
                Cell cell = row != null ? row.getCell(colIdx) : null;
                String value = CellConverter.convertToString(cell, config);
                rowArray.append(StringUtils.fromString(value));
            }
            rows.add(rowArray);
            parsedCount++;
        }

        BArray result = ValueCreator.createArrayValue(resultType);
        for (BArray row : rows) {
            result.append(row);
        }

        return result;
    }

    /**
     * Get rows as record[].
     */
    private static Object getRowsAsRecords(Environment env, Sheet sheet, XlsxConfig config, RecordType recordType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, recordType, env, isOverwritten);

        return RecordParsingUtils.parseRowsToRecords(context);
    }

    /**
     * Get rows as Row-wrapped record[] (preserves row positions).
     */
    private static Object getRowsAsRowWrappedRecords(Environment env, Sheet sheet, XlsxConfig config,
                                                      RecordType rowWrappedType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        // Extract the inner value type (e.g., Person from PersonRow)
        Type innerValueType = RowTypeUtils.extractValueType(rowWrappedType);

        // If the inner value type is a record, parse with Row wrapper support
        if (innerValueType != null && innerValueType.getTag() == TypeTags.RECORD_TYPE_TAG) {
            RecordType innerRecordType = (RecordType) innerValueType;

            RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                    sheet, usedRange, config, innerRecordType, env, isOverwritten);

            return RecordParsingUtils.parseRowsToRowWrappedRecords(context, rowWrappedType);
        }

        // Fallback: if inner type is not a record, return empty array
        ArrayType arrayType = TypeCreator.createArrayType(rowWrappedType);
        return ValueCreator.createArrayValue(arrayType);
    }

    /**
     * Get rows as map<anydata>[].
     */
    private static Object getRowsAsMaps(Environment env, Sheet sheet, XlsxConfig config, MapType mapType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, mapType, env, isOverwritten);

        return RecordParsingUtils.parseRowsToMaps(context);
    }

    /**
     * Get the native Sheet from Ballerina object.
     */
    private static Sheet getSheet(BObject sheetObj) {
        Sheet sheet = (Sheet) sheetObj.getNativeData(SHEET_NATIVE_KEY);
        if (sheet == null) {
            throw new IllegalStateException(
                    "Sheet is not properly initialized. The workbook may have been closed.");
        }
        return sheet;
    }

    // =============================================================================
    // TABLE METHODS
    // =============================================================================

    /**
     * Get a table from this sheet by name.
     *
     * @param sheetObj Ballerina Sheet object
     * @param name     Table name
     * @return Ballerina Table object or error
     */
    public static Object getTable(BObject sheetObj, BString name) {
        Sheet sheet = getSheet(sheetObj);

        if (!(sheet instanceof XSSFSheet)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        String tableName = name.getValue();

        for (XSSFTable table : xssfSheet.getTables()) {
            if (tableName.equals(table.getName()) || tableName.equals(table.getDisplayName())) {
                return createBallerinaTable(table, xssfSheet);
            }
        }

        return DiagnosticLog.tableNotFoundError(tableName, xssfSheet.getSheetName());
    }

    /**
     * Get all tables in this sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Array of Ballerina Table objects
     */
    public static BArray getTables(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);

        // Get proper Table type from module for array creation
        Type tableType = ValueCreator.createObjectValue(
                ModuleUtils.getModule(), Constants.TABLE_TYPE).getType();
        ArrayType tableArrayType = TypeCreator.createArrayType(tableType);

        if (!(sheet instanceof XSSFSheet)) {
            // Return empty array for non-XLSX sheets
            return ValueCreator.createArrayValue(tableArrayType);
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        List<XSSFTable> tables = xssfSheet.getTables();

        BArray result = ValueCreator.createArrayValue(tableArrayType);
        for (XSSFTable table : tables) {
            result.append(createBallerinaTable(table, xssfSheet));
        }

        return result;
    }

    /**
     * Create a new table with the specified range.
     *
     * @param sheetObj Ballerina Sheet object
     * @param name     Table name
     * @param range    Range (CellRange record or string)
     * @param headers  Optional headers array
     * @return Ballerina Table object or error
     */
    public static Object createTable(BObject sheetObj, BString name, Object range, Object headers) {
        Sheet sheet = getSheet(sheetObj);

        if (!(sheet instanceof XSSFSheet)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        String tableName = name.getValue();

        try {
            // Check if table name already exists in workbook
            org.apache.poi.xssf.usermodel.XSSFWorkbook workbook =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook) xssfSheet.getWorkbook();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet s =
                        (XSSFSheet) workbook.getSheetAt(i);
                for (XSSFTable t : s.getTables()) {
                    if (tableName.equals(t.getName()) || tableName.equals(t.getDisplayName())) {
                        return DiagnosticLog.error("Table '" + tableName + "' already exists in workbook");
                    }
                }
            }

            // Parse range
            AreaReference areaRef;
            if (range instanceof BString) {
                String rangeStr = ((BString) range).getValue();
                areaRef = new AreaReference(rangeStr,
                        SpreadsheetVersion.EXCEL2007);
            } else if (range instanceof BMap) {
                @SuppressWarnings("unchecked")
                BMap<BString, Object> cellRange = (BMap<BString, Object>) range;
                int firstRow = ((Long) cellRange.get(StringUtils.fromString("firstRowIndex"))).intValue();
                int lastRow = ((Long) cellRange.get(StringUtils.fromString("lastRowIndex"))).intValue();
                int firstCol = ((Long) cellRange.get(StringUtils.fromString("firstColumnIndex"))).intValue();
                int lastCol = ((Long) cellRange.get(StringUtils.fromString("lastColumnIndex"))).intValue();

                CellReference topLeft = new CellReference(firstRow, firstCol);
                CellReference bottomRight = new CellReference(lastRow, lastCol);
                areaRef = new AreaReference(topLeft, bottomRight,
                        SpreadsheetVersion.EXCEL2007);
            } else {
                return DiagnosticLog.invalidTableRangeError("Invalid range type");
            }

            // Check for overlap with existing tables
            for (XSSFTable existing : xssfSheet.getTables()) {
                AreaReference existingArea = new AreaReference(
                        existing.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
                if (rangesOverlap(areaRef, existingArea)) {
                    return DiagnosticLog.tableOverlapError(tableName, existing.getName(), xssfSheet.getSheetName());
                }
            }

            // Create the table
            XSSFTable table = xssfSheet.createTable(areaRef);
            table.setName(tableName);
            table.setDisplayName(tableName);

            // Set headers if provided
            if (headers instanceof BArray) {
                BArray headerArray = (BArray) headers;
                List<XSSFTableColumn> columns = table.getColumns();
                for (int i = 0; i < headerArray.getLength() && i < columns.size(); i++) {
                    String headerValue = headerArray.get(i).toString();
                    columns.get(i).setName(headerValue);

                    // Also write to the cell
                    int headerRowIdx = areaRef.getFirstCell().getRow();
                    int colIdx = areaRef.getFirstCell().getCol() + i;
                    Row headerRow = xssfSheet.getRow(headerRowIdx);
                    if (headerRow == null) {
                        headerRow = xssfSheet.createRow(headerRowIdx);
                    }
                    Cell cell = headerRow.getCell(colIdx);
                    if (cell == null) {
                        cell = headerRow.createCell(colIdx);
                    }
                    cell.setCellValue(headerValue);
                }
            }

            table.updateHeaders();
            return createBallerinaTable(table, xssfSheet);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error creating table: " + e.getMessage(), e);
        }
    }

    /**
     * Create a table from data array.
     *
     * @param sheetObj         Ballerina Sheet object
     * @param name             Table name
     * @param data             Data to write
     * @param startRowIndex    Starting row
     * @param startColumnIndex Starting column
     * @return Ballerina Table object or error
     */
    public static Object createTableFromData(BObject sheetObj, BString name, BArray data,
                                              long startRowIndex, long startColumnIndex) {
        Sheet sheet = getSheet(sheetObj);

        if (!(sheet instanceof XSSFSheet)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;

        try {
            // Write data first
            BMap<BString, Object> writeOptions = ValueCreator.createMapValue();
            writeOptions.put(StringUtils.fromString("writeHeaders"), true);
            writeOptions.put(StringUtils.fromString("startRowIndex"), startRowIndex);

            Object writeResult = XlsxWriter.writeToSheet(xssfSheet, data, writeOptions);
            if (writeResult != null) {
                return writeResult; // Error from write
            }

            // Calculate the table range
            int startRow = (int) startRowIndex;
            int startCol = (int) startColumnIndex;

            // Get dimensions from data
            int rowCount = (int) data.getLength() + 1; // +1 for header
            int colCount = 1; // Default

            // Determine column count from first record
            if (data.getLength() > 0) {
                Object firstItem = data.get(0);
                if (firstItem instanceof BMap) {
                    @SuppressWarnings("unchecked")
                    BMap<BString, Object> record = (BMap<BString, Object>) firstItem;
                    colCount = record.getKeys().length;
                } else if (firstItem instanceof BArray) {
                    colCount = (int) ((BArray) firstItem).getLength();
                }
            }

            int lastRow = startRow + rowCount - 1;
            int lastCol = startCol + colCount - 1;

            CellReference topLeft = new CellReference(startRow, startCol);
            CellReference bottomRight = new CellReference(lastRow, lastCol);
            AreaReference areaRef = new AreaReference(
                    topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);

            // Create table from the written range
            return createTable(sheetObj, name, StringUtils.fromString(areaRef.formatAsString()), null);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error creating table from data: " + e.getMessage(), e);
        }
    }

    /**
     * Delete a table from this sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @param name     Table name
     * @return null on success, error if not found
     */
    public static Object deleteTable(BObject sheetObj, BString name) {
        Sheet sheet = getSheet(sheetObj);

        if (!(sheet instanceof XSSFSheet)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        String tableName = name.getValue();

        for (XSSFTable table : xssfSheet.getTables()) {
            if (tableName.equals(table.getName()) || tableName.equals(table.getDisplayName())) {
                xssfSheet.removeTable(table);
                return null;
            }
        }

        return DiagnosticLog.tableNotFoundError(tableName, xssfSheet.getSheetName());
    }

    /**
     * Create a Ballerina Table object from POI XSSFTable.
     */
    private static BObject createBallerinaTable(XSSFTable table,
                                                 XSSFSheet sheet) {
        BObject tableObj = ValueCreator.createObjectValue(ModuleUtils.getModule(), Constants.TABLE_TYPE);
        TableHandle.initTable(tableObj, table, sheet);
        return tableObj;
    }

    /**
     * Check if two area references overlap.
     */
    private static boolean rangesOverlap(AreaReference a,
                                          AreaReference b) {
        int aFirstRow = a.getFirstCell().getRow();
        int aLastRow = a.getLastCell().getRow();
        int aFirstCol = a.getFirstCell().getCol();
        int aLastCol = a.getLastCell().getCol();

        int bFirstRow = b.getFirstCell().getRow();
        int bLastRow = b.getLastCell().getRow();
        int bFirstCol = b.getFirstCell().getCol();
        int bLastCol = b.getLastCell().getCol();

        // Check if ranges overlap
        boolean rowsOverlap = aFirstRow <= bLastRow && aLastRow >= bFirstRow;
        boolean colsOverlap = aFirstCol <= bLastCol && aLastCol >= bFirstCol;

        return rowsOverlap && colsOverlap;
    }
}
