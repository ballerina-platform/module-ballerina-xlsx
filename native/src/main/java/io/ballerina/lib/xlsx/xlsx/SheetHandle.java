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

package io.ballerina.lib.xlsx.xlsx;

import io.ballerina.lib.xlsx.utils.AnnotationUtils;
import io.ballerina.lib.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.xlsx.utils.ModuleUtils;
import io.ballerina.lib.xlsx.utils.RecordParsingUtils;
import io.ballerina.lib.xlsx.utils.UsedRangeDetector;
import io.ballerina.lib.xlsx.utils.XlsxConfig;
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
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Native handle for Apache POI Sheet.
 * This class wraps the POI Sheet and provides methods called from Ballerina.
 */
public final class SheetHandle {

    // Ballerina type name for the public `Sheet` object type, whose implementation is the
    // non-public `SheetImpl` class. Native instance construction must target the concrete class name.
    public static final String SHEET_TYPE = "SheetImpl";

    // Package-private so WorkbookHandle can null this slot during close()/deleteSheet() invalidation.
    static final String SHEET_NATIVE_KEY = "sheetNative";
    static final String PARENT_WORKBOOK_KEY = "parentWorkbook";

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
    public static Object getName(BObject sheetObj) {
        try {
            Sheet sheet = getSheet(sheetObj);
            return StringUtils.fromString(sheet.getSheetName());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the used range of the sheet in A1 notation.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Used range string (e.g., "A1:D50")
     */
    public static Object getUsedRange(BObject sheetObj) {
        try {
            Sheet sheet = getSheet(sheetObj);
            CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
            return StringUtils.fromString(UsedRangeDetector.toA1Notation(range));
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the used cell range of the sheet as a structured record.
     *
     * @param sheetObj Ballerina Sheet object
     * @return CellRange record with 0-based indices, or null if sheet is empty
     */
    public static Object getUsedCellRange(BObject sheetObj) {
        try {
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
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the row count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of rows with data
     */
    public static Object getRowCount(BObject sheetObj) {
        try {
            Sheet sheet = getSheet(sheetObj);
            CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
            return (long) UsedRangeDetector.getRowCount(range);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the column count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of columns with data
     */
    public static Object getColumnCount(BObject sheetObj) {
        try {
            Sheet sheet = getSheet(sheetObj);
            CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
            return (long) UsedRangeDetector.getColumnCount(range);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
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
        try {
            Sheet sheet = getSheet(sheetObj);
            XlsxConfig config = XlsxConfig.fromParseOptions(options);
            RecordParsingUtils.validateReadConfig(config);

            // Under the typedesc<Row> signature, the describing type IS the row element type
            // (the function returns `t[]`). Dispatch directly on the row shape.
            Type describingType = TypeUtils.getReferredType(targetType.getDescribingType());
            int typeTag = describingType.getTag();

            // string[] row → string[][]
            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) describingType;
                Type elementType = TypeUtils.getReferredType(arrayType.getElementType());
                if (elementType.getTag() == TypeTags.STRING_TAG) {
                    return getRowsAsStringArray(sheet, config);
                }
            }

            // record{} → record[]
            if (typeTag == TypeTags.RECORD_TYPE_TAG) {
                return getRowsAsRecords(env, sheet, config, (RecordType) describingType);
            }

            // map<CellValue?> → map<CellValue?>[]
            if (typeTag == TypeTags.MAP_TAG) {
                return getRowsAsMaps(env, sheet, config, (MapType) describingType);
            }

            // Row (the public union) — default to string[][].
            if (typeTag == TypeTags.UNION_TAG) {
                return getRowsAsStringArray(sheet, config);
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
        try {
            Sheet sheet = getSheet(sheetObj);
            XlsxConfig config = XlsxConfig.fromParseOptions(options);
            RecordParsingUtils.validateReadConfig(config);

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
                return getRowAsRecord(sheet, row, actualRowIndex, usedRange, config, (RecordType) resolvedType);
            }

            // Row (the public union) — default to string[].
            if (typeTag == TypeTags.UNION_TAG) {
                return getRowAsStringArray(row, usedRange, config);
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
     * Get a single row as a record via the shared per-row binder.
     *
     * <p>Single-row reads are fail-fast: a fail-safe context would skip the one row the
     * caller asked for, leaving nothing to return. Passing {@code env = null} makes the
     * binder surface a typed error on a bad cell or failed constraint instead of skipping.
     * Constraint validation still applies (records always validate when enabled).</p>
     *
     * @return BMap on success, or a typed BError on a conversion / constraint / projection failure
     */
    private static Object getRowAsRecord(Sheet sheet, Row row, int rowIdx, CellRangeAddress usedRange,
                                         XlsxConfig config, RecordType recordType) {
        Map<String, Integer> headerMap;
        if (config.hasHeaders()) {
            Integer headerRowIndex = config.getHeaderRowIndex();
            Row headerRow = sheet.getRow(headerRowIndex);
            if (headerRow == null) {
                throw new BallerinaErrorException(
                        DiagnosticLog.parseError("Header row " + headerRowIndex + " is empty"));
            }
            headerMap = RecordParsingUtils.buildHeaderMap(headerRow, usedRange, config.isCaseInsensitiveHeaders());
        } else {
            headerMap = RecordParsingUtils.buildHeaderMapFromIndices(usedRange, config.isCaseInsensitiveHeaders());
        }

        RecordParsingUtils.RecordBinding binding =
                RecordParsingUtils.buildRecordBinding(recordType, headerMap, sheet.getSheetName(), config);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, recordType, null, new AtomicBoolean(false));

        return RecordParsingUtils.parseRowToRecord(row, rowIdx, binding.columnToField, binding.absentFields,
                binding.extraColumns, binding.restFieldType, context);
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
        try {
            Sheet sheet = getSheet(sheetObj);
            return XlsxWriter.writeToSheet(sheet, data, options);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
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
     * Get rows as map<CellValue?>[].
     */
    private static Object getRowsAsMaps(Environment env, Sheet sheet, XlsxConfig config, MapType mapType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, mapType, env, isOverwritten);

        return RecordParsingUtils.parseRowsToMaps(context);
    }

    /**
     * Resolve a column reference (string header name or 0-based int index) against the
     * sheet's header row at the configured header-row index. Returns the column index.
     * Throws BallerinaErrorException with a clear message if the reference cannot be resolved.
     */
    private static int resolveColumnRef(Sheet sheet, Object columnRef, XlsxConfig config) {
        if (columnRef instanceof Long) {
            return ((Long) columnRef).intValue();
        }
        if (columnRef instanceof BString) {
            String headerName = ((BString) columnRef).getValue();
            int headerRowIdx = config.hasHeaders() ? config.getHeaderRowIndex() : 0;
            Map<String, Integer> headers = XlsxWriter.existingHeaderMap(sheet, headerRowIdx);
            if (headers == null) {
                throw new BallerinaErrorException(DiagnosticLog.error(
                        "Cannot resolve column '" + headerName
                        + "': sheet has no header row at index " + headerRowIdx));
            }
            Integer col = headers.get(headerName);
            if (col == null && config.isCaseInsensitiveHeaders()) {
                // Honour caseInsensitiveHeaders for column lookup too: existingHeaderMap is
                // case-sensitive, so fall back to a case-insensitive scan.
                for (Map.Entry<String, Integer> entry : headers.entrySet()) {
                    if (entry.getKey().equalsIgnoreCase(headerName)) {
                        col = entry.getValue();
                        break;
                    }
                }
            }
            if (col == null) {
                throw new BallerinaErrorException(DiagnosticLog.error(
                        "Column header '" + headerName + "' not found in sheet"));
            }
            return col;
        }
        throw new BallerinaErrorException(DiagnosticLog.error(
                "Column reference must be a header name (string) or 0-based index (int)"));
    }

    /**
     * Resolve the effective data-start row. If the caller passed an explicit
     * {@code dataStartRowIndex}, use it; otherwise default to the row after the header row
     * (header defaults to row 0 when not specified).
     */
    private static int defaultDataStartRow(XlsxConfig config) {
        if (config.hasExplicitDataStartRowIndex()) {
            return config.getDataStartRowIndex();
        }
        int headerRowIdx = config.hasHeaders() ? config.getHeaderRowIndex() : 0;
        return headerRowIdx + 1;
    }

    /**
     * Get a column of values by header name or 0-based index.
     */
    public static Object getColumn(Environment env, BObject sheetObj, Object columnRef,
                                    BMap<BString, Object> options, BTypedesc targetType) {
        try {
            Sheet sheet = getSheet(sheetObj);
            XlsxConfig config = XlsxConfig.fromParseOptions(options);
            RecordParsingUtils.validateReadConfig(config);
            int colIdx = resolveColumnRef(sheet, columnRef, config);

            // Under typedesc<anydata>, the describing type IS the cell element type
            // (the function returns `t[]`). Synthesize the array type for the result.
            Type elementType = TypeUtils.getReferredType(targetType.getDescribingType());
            ArrayType arrType = TypeCreator.createArrayType(elementType);

            CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
            if (usedRange == null) {
                return ValueCreator.createArrayValue(arrType);
            }

            int dataStartRow = defaultDataStartRow(config);
            int endRow = usedRange.getLastRow();
            Integer rowCountLimit = config.getRowCount();

            BArray result = ValueCreator.createArrayValue(arrType);
            int parsedCount = 0;
            for (int r = dataStartRow; r <= endRow; r++) {
                if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                    break;
                }
                Row row = sheet.getRow(r);
                Cell cell = row != null ? row.getCell(colIdx) : null;
                Object value;
                try {
                    value = CellConverter.convert(cell, elementType, config);
                } catch (TypeConversionException e) {
                    String cellAddress = RecordParsingUtils.getCellAddress(colIdx, r);
                    throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                            e.getMessage(), cellAddress, r, colIdx));
                }
                result.append(value);
                parsedCount++;
            }
            return result;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting column: " + e.getMessage(), e);
        }
    }

    /**
     * Read a single cell value, bound to the target type.
     *
     * <p>The target type drives binding exactly as elsewhere: a broad target (the default
     * {@code CellValue?}) yields the cell's natural value (whole number → int, fractional →
     * decimal, date/time → ISO string), while a pinned {@code time:Civil} / {@code time:Date}
     * / {@code time:TimeOfDay} / scalar yields that type. A blank cell binds to {@code ()} for
     * a nilable target, or surfaces a typed error for a non-nilable one (same rule as a
     * required non-nilable record field over a blank cell).</p>
     */
    public static Object getCell(BObject sheetObj, long rowIdx, long colIdx, BTypedesc targetType) {
        try {
            Sheet sheet = getSheet(sheetObj);
            Type elementType = TypeUtils.getReferredType(targetType.getDescribingType());
            Row row = sheet.getRow((int) rowIdx);
            Cell cell = row != null ? row.getCell((int) colIdx) : null;

            if (cell == null) {
                // Blank cell: nil for a nilable target, otherwise a typed error.
                if (RecordParsingUtils.isNilableType(elementType)) {
                    return null;
                }
                String cellAddress = RecordParsingUtils.getCellAddress((int) colIdx, (int) rowIdx);
                return DiagnosticLog.typeConversionError(
                        "Blank cell cannot bind to non-nilable type", cellAddress, (int) rowIdx, (int) colIdx);
            }

            XlsxConfig config = XlsxConfig.fromParseOptions(ValueCreator.createMapValue());
            try {
                return CellConverter.convert(cell, elementType, config);
            } catch (TypeConversionException e) {
                String cellAddress = RecordParsingUtils.getCellAddress((int) colIdx, (int) rowIdx);
                return DiagnosticLog.typeConversionError(e.getMessage(), cellAddress,
                        (int) rowIdx, (int) colIdx);
            }
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting cell: " + e.getMessage(), e);
        }
    }

    /**
     * Write a single row at the specified row index.
     *
     * <p>For string[] data, values are written positionally starting at column 0.
     * For record/map data, the sheet must have a header row at the configured
     * header-row index; each value is placed in the column matching its key/field name.
     * Records honor {@code @xlsx:Name} annotations for the header lookup.</p>
     */
    public static Object setRow(BObject sheetObj, long rowIdx, Object rowData,
                                 BMap<BString, Object> options) {
        try {
            Sheet sheet = getSheet(sheetObj);
            XlsxConfig config = XlsxConfig.fromWriteOptions(options);
            int targetRow = (int) rowIdx;
            StyleCache styleCache = new StyleCache(sheet.getWorkbook());
            if (rowData instanceof BArray) {
                BArray arr = (BArray) rowData;
                Row row = sheet.getRow(targetRow);
                if (row == null) {
                    row = sheet.createRow(targetRow);
                }
                for (int j = 0; j < arr.size(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) {
                        cell = row.createCell(j);
                    }
                    CellConverter.setCellValue(cell, arr.get(j), styleCache);
                }
                return null;
            }
            if (rowData instanceof BMap) {
                int headerRowIdx = config.hasHeaders() ? config.getHeaderRowIndex() : 0;
                Map<String, Integer> headers = XlsxWriter.existingHeaderMap(sheet, headerRowIdx);
                if (headers == null) {
                    return DiagnosticLog.error(
                            "setRow with a record or map requires an existing header row at index "
                            + headerRowIdx);
                }
                @SuppressWarnings("unchecked")
                BMap<BString, Object> map = (BMap<BString, Object>) rowData;
                // Pick a record type when available (so @xlsx:Name resolution applies). When the
                // declared static type is the `Row` union, the runtime value's type is the concrete
                // narrowed member — but if it ever surfaces as the union itself, fall back to the
                // map-keys path.
                Type valueType = TypeUtils.getReferredType(TypeUtils.getType(map));
                RecordType recordType = null;
                if (valueType.getTag() == TypeTags.RECORD_TYPE_TAG) {
                    recordType = (RecordType) valueType;
                }

                Row row = sheet.getRow(targetRow);
                if (row == null) {
                    row = sheet.createRow(targetRow);
                }

                if (recordType != null) {
                    for (String fieldName : recordType.getFields().keySet()) {
                        String headerName = AnnotationUtils.getHeaderName(recordType, fieldName);
                        Integer col = headers.get(headerName);
                        if (col == null) {
                            return DiagnosticLog.error(
                                    "Field '" + fieldName + "' (header '" + headerName
                                    + "') has no matching column in the sheet header row");
                        }
                        Object value = map.get(StringUtils.fromString(fieldName));
                        Cell cell = row.getCell(col);
                        if (cell == null) {
                            cell = row.createCell(col);
                        }
                        CellConverter.setCellValue(cell, value, styleCache);
                    }
                } else {
                    // Map<anydata> or unnarrowed Row union — write by map keys.
                    for (BString key : map.getKeys()) {
                        String headerName = key.getValue();
                        Integer col = headers.get(headerName);
                        if (col == null) {
                            return DiagnosticLog.error(
                                    "Key '" + headerName
                                    + "' has no matching column in the sheet header row");
                        }
                        Cell cell = row.getCell(col);
                        if (cell == null) {
                            cell = row.createCell(col);
                        }
                        CellConverter.setCellValue(cell, map.get(key), styleCache);
                    }
                }
                return null;
            }
            return DiagnosticLog.error(
                    "setRow data must be string[], record, or map<CellValue?>");
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting row: " + e.getMessage(), e);
        }
    }

    /**
     * Write a column of values by header name or 0-based index.
     */
    public static Object setColumn(BObject sheetObj, Object columnRef, BArray data) {
        try {
            Sheet sheet = getSheet(sheetObj);
            XlsxConfig config = XlsxConfig.fromParseOptions(ValueCreator.createMapValue());
            StyleCache styleCache = new StyleCache(sheet.getWorkbook());
            int colIdx = resolveColumnRef(sheet, columnRef, config);
            int startRow = defaultDataStartRow(config);
            for (int i = 0; i < data.size(); i++) {
                int rowIdx = startRow + i;
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    row = sheet.createRow(rowIdx);
                }
                Cell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }
                CellConverter.setCellValue(cell, data.get(i), styleCache);
            }
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting column: " + e.getMessage(), e);
        }
    }

    /**
     * Internal helper for setCell, called from setCellByAddress to skip re-fetching the sheet.
     */
    private static Object setCellInternal(Sheet sheet, long rowIdx, long colIdx, Object value) {
        Row row = sheet.getRow((int) rowIdx);
        if (row == null) {
            row = sheet.createRow((int) rowIdx);
        }
        Cell cell = row.getCell((int) colIdx);
        if (cell == null) {
            cell = row.createCell((int) colIdx);
        }
        StyleCache styleCache = new StyleCache(sheet.getWorkbook());
        CellConverter.setCellValue(cell, value, styleCache);
        return null;
    }

    /**
     * Write a single cell value by 0-based row and column index.
     */
    public static Object setCell(BObject sheetObj, long rowIdx, long colIdx, Object value) {
        try {
            Sheet sheet = getSheet(sheetObj);
            return setCellInternal(sheet, rowIdx, colIdx, value);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting cell: " + e.getMessage(), e);
        }
    }

    /**
     * Write a single cell value by A1-notation address.
     */
    public static Object setCellByAddress(BObject sheetObj, BString cellAddress, Object value) {
        try {
            Sheet sheet = getSheet(sheetObj);
            CellReference ref;
            try {
                ref = new CellReference(cellAddress.getValue());
            } catch (IllegalArgumentException e) {
                return DiagnosticLog.error(
                        "Invalid cell address '" + cellAddress.getValue() + "': " + e.getMessage());
            }
            int rowIdx = ref.getRow();
            int colIdx = ref.getCol();
            if (rowIdx < 0 || colIdx < 0) {
                return DiagnosticLog.error(
                        "Invalid cell address '" + cellAddress.getValue() + "'");
            }
            return setCellInternal(sheet, rowIdx, colIdx, value);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting cell by address: " + e.getMessage(), e);
        }
    }

    /**
     * Delete a row from the sheet. Subsequent rows shift up by one.
     */
    public static Object deleteRow(BObject sheetObj, long index) {
        try {
            Sheet sheet = getSheet(sheetObj);
            int idx = (int) index;
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum < 0 || idx < 0 || idx > lastRowNum) {
                return DiagnosticLog.error(
                        "Row index " + idx + " is out of range for deletion");
            }
            Row row = sheet.getRow(idx);
            if (row != null) {
                sheet.removeRow(row);
            }
            if (idx < lastRowNum) {
                sheet.shiftRows(idx + 1, lastRowNum, -1);
            }
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error deleting row: " + e.getMessage(), e);
        }
    }

    /**
     * Rename the sheet.
     */
    public static Object rename(BObject sheetObj, BString newName) {
        try {
            String name = newName.getValue();
            WorkbookHandle.validateSheetName(name);
            Sheet sheet = getSheet(sheetObj);
            org.apache.poi.ss.usermodel.Workbook workbook = sheet.getWorkbook();
            int idx = workbook.getSheetIndex(sheet);
            if (idx < 0) {
                return DiagnosticLog.error("Sheet not found in its parent workbook");
            }
            // Refuse to rename to a name another sheet already uses (case-insensitive).
            int existing = WorkbookHandle.findSheetIndexCaseInsensitive(workbook, name);
            if (existing != -1 && existing != idx) {
                return DiagnosticLog.error("Sheet '" + name + "' already exists");
            }
            workbook.setSheetName(idx, name);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (IllegalArgumentException e) {
            return DiagnosticLog.error(
                    "Invalid sheet name '" + newName.getValue() + "': " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error renaming sheet: " + e.getMessage(), e);
        }
    }

    /**
     * Get the native Sheet from Ballerina object. Throws a {@link BallerinaErrorException}
     * carrying a typed {@code xlsx:Error} if the handle has been invalidated (workbook
     * closed or sheet deleted). Callers that wrap in the standard
     * {@code catch (BallerinaErrorException e) { return e.getBError(); }} pattern will
     * surface it cleanly to Ballerina.
     */
    private static Sheet getSheet(BObject sheetObj) {
        Sheet sheet = (Sheet) sheetObj.getNativeData(SHEET_NATIVE_KEY);
        if (sheet == null) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Sheet handle is no longer valid. The workbook may have been closed "
                            + "or the sheet deleted."));
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
        try {
            Sheet sheet = getSheet(sheetObj);

            if (!(sheet instanceof XSSFSheet)) {
                return DiagnosticLog.error("Tables are only supported in XLSX format");
            }

            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            String tableName = name.getValue();

            for (XSSFTable table : xssfSheet.getTables()) {
                if (tableName.equals(table.getName()) || tableName.equals(table.getDisplayName())) {
                    return createBallerinaTable(sheetObj, table, xssfSheet);
                }
            }

            return DiagnosticLog.tableNotFoundError(tableName, xssfSheet.getSheetName());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get all tables in this sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Array of Ballerina Table objects on success, BError on failure
     */
    public static Object getTables(BObject sheetObj) {
        try {
            Sheet sheet = getSheet(sheetObj);

            // Get proper Table type from module for array creation
            Type tableType = TypeUtils.getType(ValueCreator.createObjectValue(
                    ModuleUtils.getModule(), TableHandle.TABLE_TYPE));
            ArrayType tableArrayType = TypeCreator.createArrayType(tableType);

            if (!(sheet instanceof XSSFSheet)) {
                // Return empty array for non-XLSX sheets
                return ValueCreator.createArrayValue(tableArrayType);
            }

            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            List<XSSFTable> tables = xssfSheet.getTables();

            BArray result = ValueCreator.createArrayValue(tableArrayType);
            for (XSSFTable table : tables) {
                result.append(createBallerinaTable(sheetObj, table, xssfSheet));
            }

            return result;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error retrieving tables: " + e.getMessage(), e);
        }
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
        try {
            TableHandle.validateTableName(name.getValue());
            Sheet sheet = getSheet(sheetObj);

            if (!(sheet instanceof XSSFSheet)) {
                return DiagnosticLog.error("Tables are only supported in XLSX format");
            }

            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            String tableName = name.getValue();
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
            return createBallerinaTable(sheetObj, table, xssfSheet);

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
        try {
            Sheet sheet = getSheet(sheetObj);

            if (!(sheet instanceof XSSFSheet)) {
                return DiagnosticLog.error("Tables are only supported in XLSX format");
            }

            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            // Write data first
            BMap<BString, Object> writeOptions = ValueCreator.createMapValue();
            writeOptions.put(StringUtils.fromString("writeHeaders"), true);
            writeOptions.put(StringUtils.fromString("startRowIndex"), startRowIndex);
            writeOptions.put(StringUtils.fromString("startColumnIndex"), startColumnIndex);

            Object writeResult = XlsxWriter.writeToSheet(xssfSheet, data, writeOptions);
            if (writeResult != null) {
                return writeResult; // Error from write
            }

            // Calculate the table range
            int startRow = (int) startRowIndex;
            int startCol = (int) startColumnIndex;

            // Get dimensions from data. Record/map rows generate a header row, so the
            // table spans data rows + 1. A string[][] already carries its header as the
            // first row, so its height is exactly the supplied row count.
            int colCount = 1; // Default
            boolean arrayData = false;

            // Determine column count and row shape from the first row
            if (data.getLength() > 0) {
                Object firstItem = data.get(0);
                if (firstItem instanceof BMap) {
                    @SuppressWarnings("unchecked")
                    BMap<BString, Object> record = (BMap<BString, Object>) firstItem;
                    colCount = record.getKeys().length;
                } else if (firstItem instanceof BArray) {
                    arrayData = true;
                    colCount = (int) ((BArray) firstItem).getLength();
                }
            }

            int rowCount = arrayData ? (int) data.getLength() : (int) data.getLength() + 1;

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
        Sheet sheet;
        try {
            sheet = getSheet(sheetObj);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }

        if (!(sheet instanceof XSSFSheet)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        String tableName = name.getValue();

        for (XSSFTable table : xssfSheet.getTables()) {
            if (tableName.equals(table.getName()) || tableName.equals(table.getDisplayName())) {
                BObject parentWorkbook = (BObject) sheetObj.getNativeData(PARENT_WORKBOOK_KEY);
                if (parentWorkbook != null) {
                    WorkbookHandle.invalidateHandlesForTable(parentWorkbook, table);
                }
                xssfSheet.removeTable(table);
                return null;
            }
        }

        return DiagnosticLog.tableNotFoundError(tableName, xssfSheet.getSheetName());
    }

    /**
     * Create a Ballerina Table object from POI XSSFTable and register it as vended from
     * the sheet's parent workbook (so close()/deleteSheet() can invalidate it).
     */
    private static BObject createBallerinaTable(BObject sheetObj, XSSFTable table, XSSFSheet sheet) {
        BObject tableObj = ValueCreator.createObjectValue(ModuleUtils.getModule(), TableHandle.TABLE_TYPE);
        TableHandle.initTable(tableObj, table, sheet);
        BObject parentWorkbook = (BObject) sheetObj.getNativeData(PARENT_WORKBOOK_KEY);
        if (parentWorkbook != null) {
            WorkbookHandle.registerVendedHandle(parentWorkbook, tableObj);
        }
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
