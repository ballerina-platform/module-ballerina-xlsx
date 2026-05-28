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

package io.ballerina.stdlib.xlsx.xlsx;

import io.ballerina.runtime.api.types.ArrayType;
import io.ballerina.runtime.api.types.Field;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.stdlib.xlsx.utils.AnnotationUtils;
import io.ballerina.stdlib.xlsx.utils.DiagnosticLog;
import io.ballerina.stdlib.xlsx.utils.XlsxConfig;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * Core XLSX writing logic using Apache POI.
 */
public final class XlsxWriter {

    private XlsxWriter() {
        // Private constructor to prevent instantiation
    }

    /**
     * Write Ballerina data directly to an XLSX file.
     *
     * @param filePath Path to the output file
     * @param data     Ballerina array (string[][] or record[])
     * @param options  Write options
     * @return null on success, error on failure
     */
    public static Object writeToFile(String filePath, BArray data, BMap<BString, Object> options) {
        XlsxConfig config = XlsxConfig.fromWriteOptions(options);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(config.getWriteSheetName());

            int startRow = config.getStartRowIndex();
            int startCol = config.getStartColumnIndex();
            StyleCache styleCache = new StyleCache(workbook);
            Object dispatchResult = dispatchWrite(sheet, data, config, startRow, startCol, styleCache,
                    "XLSX export");
            if (dispatchResult != null) {
                return dispatchResult; // error from dispatch
            }

            // Atomic save: serialize → temp file → atomic rename. If any step fails,
            // the destination file (if it existed) is untouched.
            WorkbookHandle.writeAtomically(java.nio.file.Paths.get(filePath), workbook);
            return null; // Success

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (IOException e) {
            return DiagnosticLog.error("Failed to write XLSX file: " + e.getMessage(), e);
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing XLSX: " + e.getMessage(), e);
        }
    }

    /**
     * Dispatch a BArray write to the correct branch (array / record / map).
     *
     * <p>Returns {@code null} on success or a {@code BError} on failure (suitable for the
     * caller to return directly to Ballerina).</p>
     *
     * <p>The declared element type drives dispatch when it's concrete (string[][], record[],
     * map[]). When the declared element type is the public {@code xlsx:Row} union — which
     * the {@code Data} alias resolves to — runtime inspection of {@code data.get(0)} picks
     * the path: a {@link BArray} element triggers the array writer, a record-typed
     * {@link BMap} triggers the record writer, and any other {@link BMap} triggers the
     * map writer. An empty array is a no-op.</p>
     */
    static Object dispatchWrite(Sheet sheet, BArray data, XlsxConfig config, int startRow,
                                 int startCol, StyleCache styleCache, String contextLabel) {
        // Unwrap any type reference on the array itself (e.g., when contextually typed as `Data`,
        // data.getType() returns a BTypeReferenceType, not the underlying ArrayType).
        Type dataType = TypeUtils.getReferredType(data.getType());
        Type elementType = ((ArrayType) dataType).getElementType();
        Type resolvedElementType = TypeUtils.getReferredType(elementType);
        int elementTag = resolvedElementType.getTag();

        if (elementTag == TypeTags.ARRAY_TAG) {
            writeArrayData(sheet, data, startRow, startCol, styleCache);
            return null;
        }
        if (elementTag == TypeTags.RECORD_TYPE_TAG) {
            writeRecordData(sheet, data, (RecordType) resolvedElementType, config, startRow, startCol, styleCache);
            return null;
        }
        if (elementTag == TypeTags.MAP_TAG) {
            writeMapData(sheet, data, config, startRow, startCol, styleCache);
            return null;
        }
        if (elementTag == TypeTags.UNION_TAG) {
            // Public `Row` union — pick a path from the first element's runtime type.
            if (data.size() == 0) {
                return null; // no-op for empty input
            }
            Object first = data.get(0);
            if (first instanceof BArray) {
                writeArrayData(sheet, data, startRow, startCol, styleCache);
                return null;
            }
            if (first instanceof BMap) {
                Type firstType = TypeUtils.getReferredType(TypeUtils.getType(first));
                if (firstType.getTag() == TypeTags.RECORD_TYPE_TAG) {
                    writeRecordData(sheet, data, (RecordType) firstType, config, startRow, startCol, styleCache);
                } else {
                    writeMapData(sheet, data, config, startRow, startCol, styleCache);
                }
                return null;
            }
            return DiagnosticLog.error(
                    "Unsupported row element runtime type for " + contextLabel + ": " + first);
        }
        return DiagnosticLog.error(
                "Unsupported data type for " + contextLabel + ": " + resolvedElementType);
    }

    private static void writeArrayData(Sheet sheet, BArray data, int startRow, int startCol,
                                        StyleCache styleCache) {
        for (int i = 0; i < data.size(); i++) {
            BArray rowData = (BArray) data.get(i);
            Row row = sheet.createRow(startRow + i);

            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(startCol + j);
                Object value = rowData.get(j);
                CellConverter.setCellValue(cell, value, styleCache);
            }
        }
    }

    /**
     * Read an existing header row at the given index and return a header-name → column-index map.
     * Returns {@code null} if the row is absent or has no string-valued cells — signal to caller
     * that the sheet has no headers and sequential positional write should be used.
     */
    static Map<String, Integer> existingHeaderMap(Sheet sheet, int headerRowIdx) {
        Row headerRow = sheet.getRow(headerRowIdx);
        if (headerRow == null) {
            return null;
        }
        Map<String, Integer> map = new HashMap<>();
        short last = headerRow.getLastCellNum();
        for (int c = 0; c < last; c++) {
            Cell cell = headerRow.getCell(c);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String header = cell.getStringCellValue();
                if (header != null && !header.isEmpty()) {
                    map.put(header, c);
                }
            }
        }
        return map.isEmpty() ? null : map;
    }

    /**
     * Write record[] data to sheet.
     *
     * <p>Dispatches on header presence at {@code startRow}:</p>
     * <ul>
     *   <li>If an existing header row is detected, each record field is resolved to its
     *       target column via {@code @xlsx:Name} (or field name) against the existing headers.
     *       Unrelated columns in the target rows are preserved. An unmatched field raises a
     *       {@link BallerinaErrorException} (silent column shifting).</li>
     *   <li>Otherwise (fresh sheet), uses sequential positional write — emits headers per the
     *       record's field order, then data rows in matching order.</li>
     * </ul>
     */
    private static void writeRecordData(Sheet sheet, BArray data, RecordType recordType,
                                         XlsxConfig config, int startRow, int startCol,
                                         StyleCache styleCache) {
        if (data.size() == 0) {
            return;
        }

        Map<String, Field> fields = recordType.getFields();
        List<String> fieldNames = new ArrayList<>(fields.keySet());

        // Build header names using @xlsx:Name annotations where present
        List<String> headerNames = new ArrayList<>();
        for (String fieldName : fieldNames) {
            headerNames.add(AnnotationUtils.getHeaderName(recordType, fieldName));
        }

        Map<String, Integer> existingHeaders = existingHeaderMap(sheet, startRow);

        if (existingHeaders != null) {
            // Header-position write: resolve each field's column from existing headers.
            // Preserves unrelated columns in target rows.
            int[] resolvedCols = new int[fieldNames.size()];
            for (int j = 0; j < fieldNames.size(); j++) {
                Integer col = existingHeaders.get(headerNames.get(j));
                if (col == null) {
                    throw new BallerinaErrorException(DiagnosticLog.error(
                            "Field '" + fieldNames.get(j) + "' (header '" + headerNames.get(j)
                            + "') has no matching column in the existing sheet header row"));
                }
                resolvedCols[j] = col;
            }

            // Data starts on the row after the existing header row
            int dataStartRow = startRow + 1;

            for (int i = 0; i < data.size(); i++) {
                Object item = data.get(i);
                if (!(item instanceof BMap)) {
                    throw new BallerinaErrorException(DiagnosticLog.error(
                            "Row " + i + " has an incompatible shape for record write. "
                                    + "All rows in a record[] write must be records."));
                }
                @SuppressWarnings("unchecked")
                BMap<BString, Object> record = (BMap<BString, Object>) item;
                Row row = sheet.getRow(dataStartRow + i);
                if (row == null) {
                    row = sheet.createRow(dataStartRow + i);
                }
                for (int j = 0; j < fieldNames.size(); j++) {
                    Cell cell = row.createCell(resolvedCols[j]);
                    String fieldName = fieldNames.get(j);
                    Object value = record.get(io.ballerina.runtime.api.utils.StringUtils.fromString(fieldName));
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            }
        } else {
            // Fresh-sheet write: sequential positional layout
            int currentRow = startRow;

            if (config.isWriteHeaders()) {
                Row headerRow = sheet.createRow(currentRow++);
                for (int i = 0; i < headerNames.size(); i++) {
                    Cell cell = headerRow.createCell(startCol + i);
                    cell.setCellValue(headerNames.get(i));
                }
            }

            for (int i = 0; i < data.size(); i++) {
                Object item = data.get(i);
                if (!(item instanceof BMap)) {
                    throw new BallerinaErrorException(DiagnosticLog.error(
                            "Row " + i + " has an incompatible shape for record write. "
                                    + "All rows in a record[] write must be records."));
                }
                @SuppressWarnings("unchecked")
                BMap<BString, Object> record = (BMap<BString, Object>) item;
                Row row = sheet.createRow(currentRow++);

                for (int j = 0; j < fieldNames.size(); j++) {
                    Cell cell = row.createCell(startCol + j);
                    String fieldName = fieldNames.get(j);
                    Object value = record.get(io.ballerina.runtime.api.utils.StringUtils.fromString(fieldName));
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            }
        }
    }

    /**
     * Write map[] data to sheet.
     *
     * <p>Dispatches on header presence at {@code startRow}: with existing headers, map keys are
     * resolved to columns by name and unrelated columns in target rows are preserved;
     * unmatched keys raise an error. Without existing headers, falls back to sequential positional
     * write using the union of all keys as headers.</p>
     */
    private static void writeMapData(Sheet sheet, BArray data, XlsxConfig config, int startRow,
                                      int startCol, StyleCache styleCache) {
        if (data.size() == 0) {
            return;
        }

        Map<String, Integer> existingHeaders = existingHeaderMap(sheet, startRow);

        if (existingHeaders != null) {
            // Header-position write: resolve each map key against existing headers
            int dataStartRow = startRow + 1;

            for (int i = 0; i < data.size(); i++) {
                @SuppressWarnings("unchecked")
                BMap<BString, Object> map = (BMap<BString, Object>) data.get(i);
                Row row = sheet.getRow(dataStartRow + i);
                if (row == null) {
                    row = sheet.createRow(dataStartRow + i);
                }
                for (BString key : map.getKeys()) {
                    String headerName = key.getValue();
                    Integer col = existingHeaders.get(headerName);
                    if (col == null) {
                        throw new BallerinaErrorException(DiagnosticLog.error(
                                "Map key '" + headerName
                                + "' has no matching column in the existing sheet header row"));
                    }
                    Cell cell = row.createCell(col);
                    Object value = map.get(key);
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            }
        } else {
            // Fresh-sheet write: union of all keys as headers, sequential positional layout
            LinkedHashSet<String> headerSet = new LinkedHashSet<>();
            for (int i = 0; i < data.size(); i++) {
                @SuppressWarnings("unchecked")
                BMap<BString, Object> map = (BMap<BString, Object>) data.get(i);
                for (BString key : map.getKeys()) {
                    headerSet.add(key.getValue());
                }
            }
            List<String> headers = new ArrayList<>(headerSet);

            int currentRow = startRow;

            if (config.isWriteHeaders()) {
                Row headerRow = sheet.createRow(currentRow++);
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(startCol + i);
                    cell.setCellValue(headers.get(i));
                }
            }

            for (int i = 0; i < data.size(); i++) {
                @SuppressWarnings("unchecked")
                BMap<BString, Object> map = (BMap<BString, Object>) data.get(i);
                Row row = sheet.createRow(currentRow++);

                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.createCell(startCol + j);
                    String header = headers.get(j);
                    Object value = map.get(io.ballerina.runtime.api.utils.StringUtils.fromString(header));
                    CellConverter.setCellValue(cell, value, styleCache);
                }
            }
        }
    }

    /**
     * Write rows to a sheet (for Workbook API).
     *
     * @param sheet   The sheet to write to
     * @param data    Data to write
     * @param options Write options
     * @return null on success, error on failure
     */
    public static Object writeToSheet(Sheet sheet, BArray data, BMap<BString, Object> options) {
        XlsxConfig config = XlsxConfig.fromWriteOptions(options);

        try {
            int startRow = config.getStartRowIndex();
            int startCol = config.getStartColumnIndex();
            StyleCache styleCache = new StyleCache(sheet.getWorkbook());
            return dispatchWrite(sheet, data, config, startRow, startCol, styleCache, "sheet write");
        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to sheet: " + e.getMessage(), e);
        }
    }
}
