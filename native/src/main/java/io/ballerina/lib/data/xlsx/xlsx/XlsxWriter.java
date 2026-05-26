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

import io.ballerina.lib.data.xlsx.utils.AnnotationUtils;
import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.utils.XlsxConfig;
import io.ballerina.runtime.api.types.ArrayType;
import io.ballerina.runtime.api.types.Field;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(filePath)) {

            Sheet sheet = workbook.createSheet(config.getWriteSheetName());

            Type elementType = ((ArrayType) data.getType()).getElementType();
            // Resolve referenced types (important for module-defined types like `type X record {...}`)
            Type resolvedElementType = TypeUtils.getReferredType(elementType);
            int elementTag = resolvedElementType.getTag();

            int startRow = config.getStartRowIndex();

            if (elementTag == TypeTags.ARRAY_TAG) {
                // string[][] - write raw arrays
                writeArrayData(sheet, data, startRow);
            } else if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                RecordType recordType = (RecordType) resolvedElementType;
                writeRecordData(sheet, data, recordType, config, startRow);
            } else if (elementTag == TypeTags.MAP_TAG) {
                // map[] - write maps with headers
                writeMapData(sheet, data, config, startRow);
            } else {
                return DiagnosticLog.error("Unsupported data type for XLSX export: " + resolvedElementType);
            }

            // Write directly to file - efficient!
            workbook.write(fos);
            return null; // Success

        } catch (IOException e) {
            return DiagnosticLog.error("Failed to write XLSX file: " + e.getMessage(), e);
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing XLSX: " + e.getMessage(), e);
        }
    }

    /**
     * Write string[][] data to sheet.
     */
    private static void writeArrayData(Sheet sheet, BArray data, int startRow) {
        for (int i = 0; i < data.size(); i++) {
            BArray rowData = (BArray) data.get(i);
            Row row = sheet.createRow(startRow + i);

            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(j);
                Object value = rowData.get(j);
                CellConverter.setCellValue(cell, value);
            }
        }
    }

    /**
     * Write record[] data to sheet.
     */
    private static void writeRecordData(Sheet sheet, BArray data, RecordType recordType,
                                         XlsxConfig config, int startRow) {
        if (data.size() == 0) {
            return;
        }

        // Get field names for data access
        Map<String, Field> fields = recordType.getFields();
        List<String> fieldNames = new ArrayList<>(fields.keySet());

        // Build header names using @xlsx:Name annotations where present
        List<String> headerNames = new ArrayList<>();
        for (String fieldName : fieldNames) {
            headerNames.add(AnnotationUtils.getHeaderName(recordType, fieldName));
        }

        int currentRow = startRow;

        // Write headers if configured
        if (config.isWriteHeaders()) {
            Row headerRow = sheet.createRow(currentRow++);
            for (int i = 0; i < headerNames.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headerNames.get(i));
            }
        }

        // Write data rows
        for (int i = 0; i < data.size(); i++) {
            Object item = data.get(i);
            if (!(item instanceof BMap)) {
                continue; // Skip non-map items
            }
            @SuppressWarnings("unchecked")
            BMap<BString, Object> record = (BMap<BString, Object>) item;
            Row row = sheet.createRow(currentRow++);

            for (int j = 0; j < fieldNames.size(); j++) {
                Cell cell = row.createCell(j);
                String fieldName = fieldNames.get(j);
                Object value = record.get(io.ballerina.runtime.api.utils.StringUtils.fromString(fieldName));
                CellConverter.setCellValue(cell, value);
            }
        }
    }

    /**
     * Write map[] data to sheet.
     */
    private static void writeMapData(Sheet sheet, BArray data, XlsxConfig config, int startRow) {
        if (data.size() == 0) {
            return;
        }

        // Get all unique keys from all maps for headers
        // Using LinkedHashSet for O(1) contains/add while maintaining insertion order
        LinkedHashSet<String> headerSet = new LinkedHashSet<>();
        for (int i = 0; i < data.size(); i++) {
            @SuppressWarnings("unchecked")
            BMap<BString, Object> map = (BMap<BString, Object>) data.get(i);
            for (BString key : map.getKeys()) {
                headerSet.add(key.getValue()); // O(1), auto-dedupes
            }
        }
        List<String> headers = new ArrayList<>(headerSet);

        int currentRow = startRow;

        // Write headers if configured
        if (config.isWriteHeaders()) {
            Row headerRow = sheet.createRow(currentRow++);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }
        }

        // Write data rows
        for (int i = 0; i < data.size(); i++) {
            @SuppressWarnings("unchecked")
            BMap<BString, Object> map = (BMap<BString, Object>) data.get(i);
            Row row = sheet.createRow(currentRow++);

            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.createCell(j);
                String header = headers.get(j);
                Object value = map.get(io.ballerina.runtime.api.utils.StringUtils.fromString(header));
                CellConverter.setCellValue(cell, value);
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
            Type elementType = ((ArrayType) data.getType()).getElementType();
            // Resolve referenced types (important for module-defined types like `type X record {...}`)
            Type resolvedElementType = TypeUtils.getReferredType(elementType);
            int elementTag = resolvedElementType.getTag();

            int startRow = config.getStartRowIndex();

            if (elementTag == TypeTags.ARRAY_TAG) {
                writeArrayData(sheet, data, startRow);
            } else if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                RecordType recordType = (RecordType) resolvedElementType;
                writeRecordData(sheet, data, recordType, config, startRow);
            } else if (elementTag == TypeTags.MAP_TAG) {
                writeMapData(sheet, data, config, startRow);
            } else {
                return DiagnosticLog.error("Unsupported data type for sheet write: " + resolvedElementType);
            }

            return null; // Success

        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to sheet: " + e.getMessage(), e);
        }
    }
}
