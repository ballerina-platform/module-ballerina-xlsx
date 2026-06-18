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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Native handle for Apache POI XSSFTable.
 * This class wraps the POI Table and provides methods called from Ballerina.
 */
public final class TableHandle {

    // Ballerina type name for the public `Table` object type, whose implementation is the
    // non-public `TableImpl` class. Native instance construction must target the concrete class name.
    public static final String TABLE_TYPE = "TableImpl";

    // Package-private so WorkbookHandle/SheetHandle can null these slots during invalidation.
    static final String TABLE_NATIVE_KEY = "tableNative";
    static final String SHEET_NATIVE_KEY = "sheetNative";
    static final String PARENT_WORKBOOK_KEY = "parentWorkbook";

    private TableHandle() {
        // Private constructor to prevent instantiation
    }

    /**
     * Initialize a Ballerina Table object with a POI Table.
     *
     * @param tableObj Ballerina Table object
     * @param table    POI XSSFTable
     * @param sheet    POI Sheet containing the table
     */
    public static void initTable(BObject tableObj, XSSFTable table, XSSFSheet sheet) {
        tableObj.addNativeData(TABLE_NATIVE_KEY, table);
        tableObj.addNativeData(SHEET_NATIVE_KEY, sheet);
    }

    // === Identity Methods ===

    /**
     * Get the table name.
     *
     * @param tableObj Ballerina Table object
     * @return Table name
     */
    public static Object getName(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            return StringUtils.fromString(table.getName());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the table display name.
     *
     * @param tableObj Ballerina Table object
     * @return Table display name
     */
    public static Object getDisplayName(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            return StringUtils.fromString(table.getDisplayName());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the sheet name containing this table.
     *
     * @param tableObj Ballerina Table object
     * @return Sheet name
     */
    public static Object getSheetName(BObject tableObj) {
        try {
            XSSFSheet sheet = getSheet(tableObj);
            return StringUtils.fromString(sheet.getSheetName());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    // === Range/Dimensions Methods ===

    /**
     * Get the full table range (including headers and totals) as a CellRange record.
     *
     * @param tableObj Ballerina Table object
     * @return CellRange record, or error if the handle is invalid
     */
    public static Object getCellRange(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            return createCellRange(area);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the full table range (including headers and totals) in A1 notation.
     *
     * @param tableObj Ballerina Table object
     * @return A1-notation range string, or error if the handle is invalid
     */
    public static Object getRange(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            return StringUtils.fromString(area.formatAsString());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the data range of the table (excluding headers and totals) as a CellRange record.
     *
     * @param tableObj Ballerina Table object
     * @return CellRange record, or error if the handle is invalid
     */
    public static Object getDataCellRange(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);

            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();

            // Exclude header row (first row)
            int dataFirstRow = firstRow + 1;
            // Exclude totals row if present (use getTotalsRowCount(), not isHasTotalsRow())
            int dataLastRow = lastRow - table.getTotalsRowCount();

            BMap<BString, Object> cellRange = ValueCreator.createRecordValue(
                    ModuleUtils.getModule(), "CellRange");
            cellRange.put(StringUtils.fromString("firstRowIndex"), (long) dataFirstRow);
            cellRange.put(StringUtils.fromString("lastRowIndex"), (long) dataLastRow);
            cellRange.put(StringUtils.fromString("firstColumnIndex"), (long) firstCol);
            cellRange.put(StringUtils.fromString("lastColumnIndex"), (long) lastCol);

            return cellRange;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the data range of the table (excluding headers and totals) in A1 notation.
     *
     * @param tableObj Ballerina Table object
     * @return A1-notation range string, or error if the handle is invalid
     */
    public static Object getDataRange(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);

            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();

            int dataFirstRow = firstRow + 1;
            int dataLastRow = lastRow - table.getTotalsRowCount();

            AreaReference dataArea = new AreaReference(
                    new CellReference(dataFirstRow, firstCol),
                    new CellReference(dataLastRow, lastCol),
                    SpreadsheetVersion.EXCEL2007);
            return StringUtils.fromString(dataArea.formatAsString());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the number of data rows (excluding headers and totals).
     *
     * @param tableObj Ballerina Table object
     * @return Data row count
     */
    public static Object getRowCount(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);

            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();

            // Calculate data rows: total rows - header row - totals row (if present)
            // Use getTotalsRowCount() instead of isHasTotalsRow() for accurate count
            int totalRows = lastRow - firstRow + 1;
            int headerRows = 1;  // Always has one header row
            int totalsRows = table.getTotalsRowCount();

            return (long) (totalRows - headerRows - totalsRows);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the number of columns.
     *
     * @param tableObj Ballerina Table object
     * @return Column count
     */
    public static Object getColumnCount(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            return (long) table.getColumnCount();
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    // === Header Methods ===

    /**
     * Get the column headers.
     *
     * @param tableObj Ballerina Table object
     * @return Array of header strings
     */
    public static Object getHeaders(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            List<XSSFTableColumn> columns = table.getColumns();

            ArrayType stringArrayType = TypeCreator.createArrayType(
                    io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
            BArray headers = ValueCreator.createArrayValue(stringArrayType);

            for (XSSFTableColumn column : columns) {
                headers.append(StringUtils.fromString(column.getName()));
            }

            return headers;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    // === Data Access Methods ===

    /**
     * Get all data rows from the table.
     *
     * @param env        Ballerina runtime environment
     * @param tableObj   Ballerina Table object
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Array of rows (string[][] or record[])
     */
    public static Object getRows(Environment env, BObject tableObj, BMap<BString, Object> options,
                                  BTypedesc targetType) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);
            return parseTableInternal(env, table, sheet, options, targetType);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting table rows: " + e.getMessage(), e);
        }
    }

    /**
     * Parse a POI XSSFTable directly to the user's target shape, without going
     * through a Ballerina Table handle. Used by {@code Native.parseTable} which
     * has the XSSF objects from {@code WorkbookFactory.create} and never vends
     * a Table BObject.
     *
     * Shared dispatch with {@link #getRows(Environment, BObject, BMap, BTypedesc)} —
     * both call into {@link #parseTableInternal} below.
     */
    public static Object parseFromXSSFTable(Environment env, XSSFTable table, XSSFSheet sheet,
                                             BMap<BString, Object> options, BTypedesc targetType) {
        try {
            return parseTableInternal(env, table, sheet, options, targetType);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing table: " + e.getMessage(), e);
        }
    }

    /**
     * Shared dispatch: resolve target typedesc, route to the appropriate row-shape
     * helper. Caller wraps in the standard try/catch (BallerinaErrorException /
     * Exception) → BError conversion.
     */
    private static Object parseTableInternal(Environment env, XSSFTable table, XSSFSheet sheet,
                                              BMap<BString, Object> options, BTypedesc targetType) {
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
                return getTableRowsAsStringArray(table, sheet, config);
            }
        }

        // record{} → record[]
        if (typeTag == TypeTags.RECORD_TYPE_TAG) {
            return getTableRowsAsRecords(env, table, sheet, config, (RecordType) describingType);
        }

        // map<CellValue> → map<CellValue>[]
        if (typeTag == TypeTags.MAP_TAG) {
            return getTableRowsAsMaps(env, table, sheet, config, (MapType) describingType);
        }

        // Row (the public union) — default to string[][].
        if (typeTag == TypeTags.UNION_TAG) {
            return getTableRowsAsStringArray(table, sheet, config);
        }

        // Default: string[][]
        return getTableRowsAsStringArray(table, sheet, config);
    }

    /**
     * Get a single data row from the table by index.
     *
     * @param env        Ballerina runtime environment
     * @param tableObj   Ballerina Table object
     * @param index      Row index (0-based within data range)
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Single row (string[] or record{})
     */
    public static Object getRow(Environment env, BObject tableObj, long index, BMap<BString, Object> options,
                                 BTypedesc targetType) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);
            XlsxConfig config = XlsxConfig.fromParseOptions(options);
            RecordParsingUtils.validateReadConfig(config);

            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();

            // Data starts after header row
            int dataFirstRow = firstRow + 1;
            int dataLastRow = lastRow - table.getTotalsRowCount();

            int actualRowIndex = dataFirstRow + (int) index;

            if (actualRowIndex < dataFirstRow || actualRowIndex > dataLastRow) {
                long maxIndex = dataLastRow - dataFirstRow;
                return DiagnosticLog.error("Row index " + index + " out of range (0-" + maxIndex + ")");
            }

            Row row = sheet.getRow(actualRowIndex);

            Type describingType = targetType.getDescribingType();
            Type resolvedType = TypeUtils.getReferredType(describingType);
            int typeTag = resolvedType.getTag();

            // string[]
            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) resolvedType;
                Type elementType = TypeUtils.getReferredType(arrayType.getElementType());
                if (elementType.getTag() == TypeTags.STRING_TAG) {
                    return getRowAsStringArray(row, firstCol, lastCol, config);
                }
            }

            // record{}
            if (typeTag == TypeTags.RECORD_TYPE_TAG) {
                return getTableRowAsRecord(table, sheet, row, actualRowIndex, config, (RecordType) resolvedType);
            }

            // Row (the public union) — default to string[].
            if (typeTag == TypeTags.UNION_TAG) {
                return getRowAsStringArray(row, firstCol, lastCol, config);
            }

            // Default: string[]
            return getRowAsStringArray(row, firstCol, lastCol, config);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting table row: " + e.getMessage(), e);
        }
    }

    /**
     * Put rows into the table. By default (REPLACE) the data region is resized to fit the data
     * exactly; APPEND adds the rows below the existing data.
     *
     * @param tableObj Ballerina Table object
     * @param data     Data to write
     * @param options  Table write options (tableWriteMode)
     * @return null on success, error on failure
     */
    public static Object putRows(BObject tableObj, BArray data, BMap<BString, Object> options) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);
            XlsxConfig config = XlsxConfig.fromWriteOptions(options);
            boolean append = "APPEND".equals(config.getTableWriteMode());
            return writeTableInternal(table, sheet, data, append, config.getTableInsertAt());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to table: " + e.getMessage(), e);
        }
    }

    /**
     * Write data to a POI XSSFTable directly, without going through a Ballerina
     * Table handle. Used by {@code Native.writeTable} which has the XSSF objects
     * from {@code WorkbookFactory.create} and never vends a Table BObject.
     *
     * Shared dispatch with {@link #putRows(BObject, BArray, BMap)} — both call
     * into {@link #writeTableInternal} below.
     */
    public static Object writeToXSSFTable(XSSFTable table, XSSFSheet sheet, BArray data,
                                          BMap<BString, Object> options) {
        try {
            XlsxConfig config = XlsxConfig.fromWriteOptions(options);
            boolean append = "APPEND".equals(config.getTableWriteMode());
            return writeTableInternal(table, sheet, data, append, config.getTableInsertAt());
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to table: " + e.getMessage(), e);
        }
    }

    /**
     * Delete a single data row from the table by 0-based data-row index. The table shrinks to fit:
     * the totals row and any content below it are pulled up to close the gap. Excel requires at
     * least one data row, so the last remaining data row cannot be deleted. A shrink that would
     * shift another table is refused with a {@code TableOverlapError}.
     *
     * @param tableObj Ballerina Table object
     * @param index    0-based row index within the data range
     * @return null on success, error on failure
     */
    public static Object deleteRow(BObject tableObj, long index) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();
            int dataFirstRow = firstRow + 1;
            int totalsRowCount = table.getTotalsRowCount();
            boolean hasTotals = totalsRowCount > 0;
            int currentDataRows = lastRow - firstRow - totalsRowCount;
            int idx = (int) index;

            if (idx < 0 || idx >= currentDataRows) {
                return DiagnosticLog.invalidTableRangeError("Row index " + idx
                        + " is out of range for a table with " + currentDataRows + " data rows");
            }
            if (currentDataRows <= 1) {
                return DiagnosticLog.invalidTableRangeError("Cannot delete the only data row of table '"
                        + table.getName() + "': a table requires at least one data row");
            }

            // Deleting pulls the totals row + everything below up; another table caught in that
            // shift would have its cells moved but its definition left stale, so refuse instead.
            String colliding = RowShifter.firstTableShiftedBy(sheet, dataFirstRow + idx, table.getName());
            if (colliding != null) {
                return DiagnosticLog.tableResizeOverlapError(table.getName(), colliding, sheet.getSheetName());
            }

            RowShifter.removeRows(sheet, dataFirstRow + idx, 1);
            resizeTableArea(table, firstRow, firstCol, lastCol, currentDataRows - 1, hasTotals);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error deleting table row: " + e.getMessage(), e);
        }
    }

    /**
     * Shared write dispatch. Resizes the table to fit the incoming data — growing or shrinking
     * the data range so no stale rows survive (REPLACE), or inserting the rows below the existing
     * data (APPEND). The totals row and any content below the table ride along with the shift; a
     * resize that would shift another table is refused with a {@code TableOverlapError}.
     */
    private static Object writeTableInternal(XSSFTable table, XSSFSheet sheet, BArray data, boolean append,
                                             Integer insertAtOverride) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();
        int lastCol = area.getLastCell().getCol();

        int dataFirstRow = firstRow + 1;  // data starts after the header row
        int totalsRowCount = table.getTotalsRowCount();
        boolean hasTotals = totalsRowCount > 0;
        int currentDataRows = lastRow - firstRow - totalsRowCount;
        int n = (int) data.getLength();

        // Resolve the new data-row count, where rows are written, and where the shift happens.
        // REPLACE: the data becomes exactly n rows (>= 1, since Excel needs a data row; an empty
        //          write clears the table to a single blank row); grow/shrink at the bottom.
        // APPEND: n rows are inserted at data-row `p` (default bottom), shifting the rows from there
        //          down — existing rows are left untouched.
        int newDataRows;
        int writeStart;
        int shiftAt;
        if (append) {
            if (n == 0) {
                return null;  // nothing to append
            }
            int p = (insertAtOverride != null) ? insertAtOverride : currentDataRows;
            if (p < 0 || p > currentDataRows) {
                return DiagnosticLog.invalidTableRangeError("insertAt " + p
                        + " is out of range for a table with " + currentDataRows + " data rows");
            }
            newDataRows = currentDataRows + n;
            writeStart = dataFirstRow + p;
            shiftAt = dataFirstRow + p;
        } else {
            newDataRows = Math.max(n, 1);
            writeStart = dataFirstRow;
            shiftAt = dataFirstRow + currentDataRows;
        }

        int delta = newDataRows - currentDataRows;     // > 0 grow, < 0 shrink, 0 same size

        // Overlap pre-check, before any mutation: a nonzero resize shifts the totals row and
        // everything below it; another table caught in that shift would have its cells moved but
        // its definition left stale, so refuse rather than corrupt it.
        if (delta != 0) {
            String colliding = RowShifter.firstTableShiftedBy(sheet, shiftAt, table.getName());
            if (colliding != null) {
                return DiagnosticLog.tableResizeOverlapError(table.getName(), colliding, sheet.getSheetName());
            }
        }

        // Make room (grow) or pull rows up (shrink), carrying the totals row + content below.
        if (delta > 0) {
            RowShifter.makeRoom(sheet, shiftAt, delta);
        } else if (delta < 0) {
            RowShifter.removeRows(sheet, dataFirstRow + newDataRows, -delta);
        }

        // Resize the table area to the new data-row count (plus the totals row, if any).
        resizeTableArea(table, firstRow, firstCol, lastCol, newDataRows, hasTotals);

        // Resolve column index by table header name. Record/map rows route values to the column
        // matching the field/key name (with @xlsx:Name support for records); array rows fall back
        // to positional placement at firstCol + i.
        StyleCache styleCache = new StyleCache(sheet.getWorkbook());
        List<XSSFTableColumn> tableColumns = table.getColumns();
        Map<String, Integer> headerToCol = new HashMap<>();
        for (int i = 0; i < tableColumns.size(); i++) {
            headerToCol.put(tableColumns.get(i).getName(), firstCol + i);
        }

        // Data is written without headers — the table carries its own header row.
        for (int i = 0; i < n; i++) {
            int rowIdx = writeStart + i;
            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }
            writeRowData(row, firstCol, data.get(i), headerToCol, styleCache);
        }

        // REPLACE with empty input keeps a single blank data row — clear its cells.
        if (!append && n == 0) {
            blankRowCells(sheet, dataFirstRow, firstCol, lastCol);
        }

        return null;
    }

    /**
     * Re-set the table's area to span {@code newDataRows} data rows (plus the totals row, if any),
     * keeping the header row and column span fixed, then resync the header model. Shared by the
     * write/resize paths so the table {@code ref} always follows a row-count change.
     */
    private static void resizeTableArea(XSSFTable table, int firstRow, int firstCol, int lastCol,
                                        int newDataRows, boolean hasTotals) {
        int dataFirstRow = firstRow + 1;
        int newLastRow = dataFirstRow + newDataRows - 1 + (hasTotals ? 1 : 0);
        CellReference topLeft = new CellReference(firstRow, firstCol);
        CellReference bottomRight = new CellReference(newLastRow, lastCol);
        AreaReference newArea = new AreaReference(topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);
        table.getCTTable().setRef(newArea.formatAsString());
        table.updateHeaders();
    }

    /**
     * Whether two rectangular areas intersect (used by the resize overlap guard).
     */
    private static boolean rangesOverlap(AreaReference a, AreaReference b) {
        int aFirstRow = a.getFirstCell().getRow();
        int aLastRow = a.getLastCell().getRow();
        int aFirstCol = a.getFirstCell().getCol();
        int aLastCol = a.getLastCell().getCol();
        int bFirstRow = b.getFirstCell().getRow();
        int bLastRow = b.getLastCell().getRow();
        int bFirstCol = b.getFirstCell().getCol();
        int bLastCol = b.getLastCell().getCol();
        boolean rowsOverlap = aFirstRow <= bLastRow && aLastRow >= bFirstRow;
        boolean colsOverlap = aFirstCol <= bLastCol && aLastCol >= bFirstCol;
        return rowsOverlap && colsOverlap;
    }

    /**
     * Blank the cells of a single row across the table's column span.
     */
    private static void blankRowCells(Sheet sheet, int rowIdx, int firstCol, int lastCol) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            return;
        }
        for (int c = firstCol; c <= lastCol; c++) {
            Cell cell = row.getCell(c);
            if (cell != null) {
                cell.setBlank();
            }
        }
    }

    // === Total Row Methods ===

    /**
     * Check if the table has a total row.
     *
     * @param tableObj Ballerina Table object
     * @return true if a total row exists
     */
    public static Object hasTotalRow(BObject tableObj) {
        try {
            XSSFTable table = getTable(tableObj);
            // Use getTotalsRowCount() instead of isHasTotalsRow()
            // isHasTotalsRow() only checks "totalsRowShown" attribute which defaults to true
            // getTotalsRowCount() returns actual count of total rows (0 or 1)
            return table.getTotalsRowCount() > 0;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get the total row values.
     *
     * @param tableObj   Ballerina Table object
     * @param targetType Descriptor for the result map type ({@code map<CellValue>})
     * @return Map of column names to total values
     */
    public static Object getTotalRow(BObject tableObj, BTypedesc targetType) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);

            if (table.getTotalsRowCount() == 0) {
                return DiagnosticLog.error("Table '" + table.getName() + "' does not have a total row");
            }

            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            int totalsRowIdx = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();

            Row totalsRow = sheet.getRow(totalsRowIdx);
            List<XSSFTableColumn> columns = table.getColumns();

            // Bind the result to the declared map type (map<CellValue>) so its runtime
            // type matches the Ballerina contract rather than the wider map<anydata>.
            MapType mapType = (MapType) TypeUtils.getReferredType(targetType.getDescribingType());
            BMap<BString, Object> totals = ValueCreator.createMapValue(mapType);

            XlsxConfig defaultConfig = new XlsxConfig();

            for (int i = 0; i < columns.size(); i++) {
                XSSFTableColumn column = columns.get(i);
                String columnName = column.getName();
                int colIdx = firstCol + i;

                Cell cell = totalsRow != null ? totalsRow.getCell(colIdx) : null;
                Object value = CellConverter.convertToCellValue(cell, defaultConfig);
                totals.put(StringUtils.fromString(columnName), value);
            }

            return totals;

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error getting totals row: " + e.getMessage(), e);
        }
    }

    /**
     * Add a total row to a table and write a literal numeric total into one column.
     *
     * Package-private — reachable only from test code via a private
     * {@code @java:Method} external in {@code test_utils.bal}. The public Ballerina
     * API does not author total rows (they require formula/aggregation authoring,
     * which is out of scope); this helper exists so fixtures can contain a real
     * total row for testing the {@code getTotalRow} read path. A literal value is
     * written rather than a {@code SUM} aggregation because POI does not evaluate
     * formulas, so a cached aggregation would read back as 0.
     *
     * @param tableObj   Ballerina Table object
     * @param totalColIndex 0-based column offset (within the table) to receive the total
     * @param totalValue The literal numeric total to write into that column
     * @return null on success, error on failure
     */
    public static Object setTotalRowNative(BObject tableObj, long totalColIndex, double totalValue) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);

            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            int firstRow = area.getFirstCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();
            int lastRow = area.getLastCell().getRow();

            // Extend the table area downward by one row to hold the total row.
            int totalsRowIdx = lastRow + 1;
            CellReference topLeft = new CellReference(firstRow, firstCol);
            CellReference bottomRight = new CellReference(totalsRowIdx, lastCol);
            AreaReference newArea = new AreaReference(topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);

            CTTable ctTable = table.getCTTable();
            ctTable.setRef(newArea.formatAsString());
            ctTable.setTotalsRowCount(1);
            ctTable.setTotalsRowShown(true);
            table.updateHeaders();

            Row totalsRow = sheet.getRow(totalsRowIdx);
            if (totalsRow == null) {
                totalsRow = sheet.createRow(totalsRowIdx);
            }
            int colIdx = firstCol + (int) totalColIndex;
            Cell cell = totalsRow.getCell(colIdx);
            if (cell == null) {
                cell = totalsRow.createCell(colIdx);
            }
            cell.setCellValue(totalValue);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting total row: " + e.getMessage(), e);
        }
    }

    // === Modification Methods ===

    /**
     * Rename the table.
     *
     * @param tableObj Ballerina Table object
     * @param newName  New table name
     * @return null on success, error on failure
     */
    public static Object rename(BObject tableObj, BString newName) {
        try {
            String name = newName.getValue();
            validateTableName(name);
            XSSFTable table = getTable(tableObj);
            table.setName(name);
            table.setDisplayName(name);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error renaming table: " + e.getMessage(), e);
        }
    }

    /**
     * Resize the table to a new range.
     *
     * @param tableObj Ballerina Table object
     * @param newRange New range — a CellRange record or an A1-notation string
     * @return null on success, error on failure
     */
    public static Object resize(BObject tableObj, Object newRange) {
        try {
            XSSFTable table = getTable(tableObj);
            XSSFSheet sheet = getSheet(tableObj);

            int firstRow;
            int lastRow;
            int firstCol;
            int lastCol;
            if (newRange instanceof BString a1) {
                AreaReference area = new AreaReference(a1.getValue(), SpreadsheetVersion.EXCEL2007);
                firstRow = area.getFirstCell().getRow();
                lastRow = area.getLastCell().getRow();
                firstCol = area.getFirstCell().getCol();
                lastCol = area.getLastCell().getCol();
            } else {
                @SuppressWarnings("unchecked")
                BMap<BString, Object> rangeRecord = (BMap<BString, Object>) newRange;
                firstRow = ((Long) rangeRecord.get(StringUtils.fromString("firstRowIndex"))).intValue();
                lastRow = ((Long) rangeRecord.get(StringUtils.fromString("lastRowIndex"))).intValue();
                firstCol = ((Long) rangeRecord.get(StringUtils.fromString("firstColumnIndex"))).intValue();
                lastCol = ((Long) rangeRecord.get(StringUtils.fromString("lastColumnIndex"))).intValue();
            }

            // Validate range shape: at least one header row + one data row, and a valid column span.
            if (firstRow >= lastRow) {
                return DiagnosticLog.invalidTableRangeError(
                        "Invalid range: must have at least one header row and one data row");
            }
            if (firstCol > lastCol) {
                return DiagnosticLog.invalidTableRangeError(
                        "Invalid range: first column is after the last column");
            }
            // Validate range bounds against the spreadsheet limits.
            SpreadsheetVersion version = SpreadsheetVersion.EXCEL2007;
            if (firstRow < 0 || firstCol < 0
                    || lastRow > version.getLastRowIndex() || lastCol > version.getLastColumnIndex()) {
                return DiagnosticLog.invalidTableRangeError("Range is outside the sheet bounds ("
                        + version.getMaxRows() + " rows x " + version.getMaxColumns() + " columns)");
            }

            // Get current column info
            int currentColCount = table.getColumnCount();
            int newColCount = lastCol - firstCol + 1;

            // Update area reference
            CellReference topLeft = new CellReference(firstRow, firstCol);
            CellReference bottomRight = new CellReference(lastRow, lastCol);
            AreaReference newArea = new AreaReference(topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);

            // Refuse a resize that would overlap another table on the sheet (mirrors createTable).
            for (XSSFTable other : sheet.getTables()) {
                if (table.getName().equals(other.getName())) {
                    continue;
                }
                AreaReference otherArea = new AreaReference(other.getArea().formatAsString(),
                        SpreadsheetVersion.EXCEL2007);
                if (rangesOverlap(newArea, otherArea)) {
                    return DiagnosticLog.tableOverlapError(table.getName(), other.getName(),
                            sheet.getSheetName());
                }
            }

            CTTable ctTable = table.getCTTable();
            ctTable.setRef(newArea.formatAsString());

            // Add new columns if needed
            if (newColCount > currentColCount) {
                CTTableColumns ctColumns = ctTable.getTableColumns();
                for (int i = currentColCount; i < newColCount; i++) {
                    CTTableColumn newCol = ctColumns.addNewTableColumn();
                    newCol.setId(i + 1);  // Column IDs are 1-based

                    // Read header from sheet cell if exists
                    Row headerRow = sheet.getRow(firstRow);
                    if (headerRow != null) {
                        Cell headerCell = headerRow.getCell(firstCol + i);
                        if (headerCell != null) {
                            newCol.setName(CellConverter.convertToString(headerCell, null));
                        } else {
                            newCol.setName("Column" + (i + 1));
                        }
                    } else {
                        newCol.setName("Column" + (i + 1));
                    }
                }
                // Update the column count attribute
                ctColumns.setCount(newColCount);
            // Remove columns if shrinking
            } else if (newColCount < currentColCount) {
                CTTableColumns ctColumns = ctTable.getTableColumns();
                // Remove extra columns from the end
                for (int i = currentColCount - 1; i >= newColCount; i--) {
                    ctColumns.removeTableColumn(i);
                }
                ctColumns.setCount(newColCount);
            }

            table.updateHeaders();
            return null;

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error resizing table: " + e.getMessage(), e);
        }
    }

    // === Helper Methods ===

    /**
     * Get table rows as string[][].
     */
    private static BArray getTableRowsAsStringArray(XSSFTable table, Sheet sheet, XlsxConfig config) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();
        int lastCol = area.getLastCell().getCol();

        // Data starts after header row
        int dataFirstRow = firstRow + 1;
        int dataLastRow = lastRow - table.getTotalsRowCount();

        ArrayType stringType = TypeCreator.createArrayType(
                io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        ArrayType resultType = TypeCreator.createArrayType(stringType);

        List<BArray> rows = new ArrayList<>();

        // Apply rowCount limit if set
        Integer rowCountLimit = config.getRowCount();
        int parsedCount = 0;

        for (int rowIdx = dataFirstRow; rowIdx <= dataLastRow; rowIdx++) {
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = sheet.getRow(rowIdx);
            BArray rowArray = ValueCreator.createArrayValue(stringType);

            for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
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
     * Get table rows as record[] via the shared per-row binder. Multi-row reads honour
     * fail-safe when configured (env threaded through), and constraint validation when
     * enabled. The header map comes from the table's column definitions; the data range
     * excludes the header row and any totals row.
     */
    private static Object getTableRowsAsRecords(Environment env, XSSFTable table, Sheet sheet, XlsxConfig config,
                                                RecordType recordType) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();
        int lastCol = area.getLastCell().getCol();

        int dataFirstRow = firstRow + 1;
        int dataLastRow = lastRow - table.getTotalsRowCount();

        Map<String, Integer> headerMap = RecordParsingUtils.buildHeaderMapFromColumns(
                columnNames(table), firstCol, config.isCaseInsensitiveHeaders());
        RecordParsingUtils.RecordBinding binding =
                RecordParsingUtils.buildRecordBinding(recordType, headerMap, sheet.getSheetName(), config);

        // usedRange = the full table area. The per-row binder reads it only for fail-safe row
        // serialization (column bounds); row iteration is bounded by dataFirstRow..dataLastRow,
        // so the totals row stays excluded.
        CellRangeAddress tableArea = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, tableArea, config, recordType, env, new AtomicBoolean(false));

        return RecordParsingUtils.bindRecordRows(context, binding, dataFirstRow, dataLastRow);
    }

    /**
     * Get table rows as map&lt;CellValue&gt;[] via the shared per-row binder. Map keys are the
     * table's column names. Honours fail-safe when configured (env threaded through).
     */
    private static Object getTableRowsAsMaps(Environment env, XSSFTable table, Sheet sheet, XlsxConfig config,
                                             MapType mapType) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();
        int lastCol = area.getLastCell().getCol();

        int dataFirstRow = firstRow + 1;
        int dataLastRow = lastRow - table.getTotalsRowCount();

        List<XSSFTableColumn> columns = table.getColumns();
        Map<Integer, String> columnToHeader = new HashMap<>();
        for (int i = 0; i < columns.size(); i++) {
            columnToHeader.put(firstCol + i, columns.get(i).getName());
        }
        Type constraintType = mapType.getConstrainedType();

        CellRangeAddress tableArea = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, tableArea, config, mapType, env, new AtomicBoolean(false));

        return RecordParsingUtils.bindMapRows(context, columnToHeader, constraintType, dataFirstRow, dataLastRow);
    }

    /**
     * Collect a table's column header names in column order.
     */
    private static List<String> columnNames(XSSFTable table) {
        List<XSSFTableColumn> columns = table.getColumns();
        List<String> names = new ArrayList<>(columns.size());
        for (XSSFTableColumn column : columns) {
            names.add(column.getName());
        }
        return names;
    }

    /**
     * Get a single row as string[].
     */
    private static BArray getRowAsStringArray(Row row, int firstCol, int lastCol, XlsxConfig config) {
        ArrayType stringType = TypeCreator.createArrayType(
                io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        BArray rowArray = ValueCreator.createArrayValue(stringType);

        for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
            Cell cell = row != null ? row.getCell(colIdx) : null;
            String value = CellConverter.convertToString(cell, config);
            rowArray.append(StringUtils.fromString(value));
        }

        return rowArray;
    }

    /**
     * Get a single table row as a record via the shared per-row binder. Single-row reads
     * are fail-fast (env = null): a fail-safe context would skip the one row the caller
     * asked for. Constraint validation still applies.
     *
     * @return BMap on success, or a typed BError on a conversion / constraint / projection failure
     */
    private static Object getTableRowAsRecord(XSSFTable table, Sheet sheet, Row row, int rowIdx,
                                              XlsxConfig config, RecordType recordType) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();
        int lastCol = area.getLastCell().getCol();

        Map<String, Integer> headerMap = RecordParsingUtils.buildHeaderMapFromColumns(
                columnNames(table), firstCol, config.isCaseInsensitiveHeaders());
        RecordParsingUtils.RecordBinding binding =
                RecordParsingUtils.buildRecordBinding(recordType, headerMap, sheet.getSheetName(), config);

        CellRangeAddress tableArea = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, tableArea, config, recordType, null, new AtomicBoolean(false));

        return RecordParsingUtils.parseRowToRecord(row, rowIdx, binding.columnToField, binding.absentFields,
                binding.extraColumns, binding.restFieldType, context);
    }

    /**
     * Write row data to a POI Row, aligning each value to its target column.
     *
     * Records: iterate fields in declaration order, resolve each header via {@code @xlsx:Name}
     * (falling back to the field name), look up the column in {@code headerToCol}. Maps:
     * iterate keys and look them up. Arrays: positional placement at {@code startCol + i}
     * (arrays have no header semantics).
     *
     * An unknown header for a record field or map key surfaces a typed
     * {@link BallerinaErrorException} so the caller sees a clear "no matching column" error
     * rather than a silent misalignment.
     */
    private static void writeRowData(Row row, int startCol, Object rowData,
                                      Map<String, Integer> headerToCol,
                                      StyleCache styleCache) {
        if (rowData instanceof BMap) {
            @SuppressWarnings("unchecked")
            BMap<BString, Object> map = (BMap<BString, Object>) rowData;

            Type valueType = TypeUtils.getReferredType(TypeUtils.getType(map));
            RecordType recordType = valueType.getTag() == TypeTags.RECORD_TYPE_TAG
                    ? (RecordType) valueType : null;

            if (recordType != null) {
                for (String fieldName : recordType.getFields().keySet()) {
                    String headerName = AnnotationUtils.getHeaderName(recordType, fieldName);
                    Integer col = headerToCol.get(headerName);
                    if (col == null) {
                        throw new BallerinaErrorException(DiagnosticLog.error(
                                "Field '" + fieldName + "' (header '" + headerName
                                        + "') has no matching column in the table"));
                    }
                    Object value = map.get(StringUtils.fromString(fieldName));
                    Cell cell = row.createCell(col);
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            } else {
                for (BString key : map.getKeys()) {
                    String headerName = key.getValue();
                    Integer col = headerToCol.get(headerName);
                    if (col == null) {
                        throw new BallerinaErrorException(DiagnosticLog.error(
                                "Map key '" + headerName
                                        + "' has no matching column in the table"));
                    }
                    Cell cell = row.createCell(col);
                    Object value = map.get(key);
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            }
        } else if (rowData instanceof BArray) {
            BArray array = (BArray) rowData;
            for (int i = 0; i < array.getLength(); i++) {
                Cell cell = row.createCell(startCol + i);
                Object value = array.get(i);
                CellConverter.setCellValue(cell, value, styleCache);
            }
        }
    }

    /**
     * Create a CellRange record from an AreaReference.
     */
    private static BMap<BString, Object> createCellRange(AreaReference area) {
        BMap<BString, Object> cellRange = ValueCreator.createRecordValue(
                ModuleUtils.getModule(), "CellRange");
        cellRange.put(StringUtils.fromString("firstRowIndex"), (long) area.getFirstCell().getRow());
        cellRange.put(StringUtils.fromString("lastRowIndex"), (long) area.getLastCell().getRow());
        cellRange.put(StringUtils.fromString("firstColumnIndex"), (long) area.getFirstCell().getCol());
        cellRange.put(StringUtils.fromString("lastColumnIndex"), (long) area.getLastCell().getCol());
        return cellRange;
    }

    /**
     * Validate a table name against Excel's rules: 1-255 characters, starts with a
     * letter or underscore, no spaces. Throws {@link BallerinaErrorException} with a
     * typed Error if invalid.
     */
    static void validateTableName(String name) {
        if (name == null || name.isEmpty()) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table name cannot be empty"));
        }
        if (name.length() > 255) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table name '" + name + "' exceeds Excel's 255-character limit"));
        }
        char first = name.charAt(0);
        if (!Character.isLetter(first) && first != '_') {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table name '" + name + "' must start with a letter or underscore"));
        }
        if (name.indexOf(' ') >= 0) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table name '" + name + "' cannot contain spaces"));
        }
    }

    /**
     * Get the native XSSFTable from Ballerina object. Throws a {@link BallerinaErrorException}
     * carrying a typed {@code xlsx:Error} if the handle has been invalidated (workbook closed
     * or table deleted). Callers should wrap in the standard
     * {@code catch (BallerinaErrorException e) { return e.getBError(); }} pattern.
     */
    private static XSSFTable getTable(BObject tableObj) {
        XSSFTable table = (XSSFTable) tableObj.getNativeData(TABLE_NATIVE_KEY);
        if (table == null) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table handle is no longer valid. The workbook may have been closed "
                            + "or the table deleted."));
        }
        return table;
    }

    /**
     * Get the native XSSFSheet from Ballerina object. Same invalidation contract as
     * {@link #getTable(BObject)}.
     */
    private static XSSFSheet getSheet(BObject tableObj) {
        XSSFSheet sheet = (XSSFSheet) tableObj.getNativeData(SHEET_NATIVE_KEY);
        if (sheet == null) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Table handle is no longer valid. The workbook may have been closed "
                            + "or the table deleted."));
        }
        return sheet;
    }
}
