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

import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.utils.RecordParsingUtils;
import io.ballerina.lib.data.xlsx.utils.RecordParsingUtils.FieldMapping;
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
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Parser for Excel Tables in XLSX files.
 * Used by the simple table API functions parseTable() and writeTable().
 */
public final class TableParser {

    private TableParser() {
        // Private constructor to prevent instantiation
    }

    /**
     * Parse data from an Excel table.
     *
     * @param env        Ballerina environment
     * @param table      The XSSFTable to parse
     * @param sheet      The sheet containing the table
     * @param options    Parse options
     * @param targetType Target type descriptor
     * @return Parsed data as BArray or BError
     */
    public static Object parseTable(Environment env, XSSFTable table, XSSFSheet sheet,
                                     BMap<BString, Object> options, BTypedesc targetType) {
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try {
            Type describingType = targetType.getDescribingType();
            int typeTag = describingType.getTag();

            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) describingType;
                Type elementType = arrayType.getElementType();
                Type resolvedElementType = TypeUtils.getReferredType(elementType);
                int elementTag = resolvedElementType.getTag();

                // string[][]
                if (elementTag == TypeTags.ARRAY_TAG) {
                    ArrayType innerArrayType = (ArrayType) resolvedElementType;
                    Type innerElementType = TypeUtils.getReferredType(innerArrayType.getElementType());
                    if (innerElementType.getTag() == TypeTags.STRING_TAG) {
                        return parseTableToStringArray(table, sheet, config);
                    }
                }

                // record[]
                if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                    return parseTableToRecords(table, sheet, config, (RecordType) resolvedElementType);
                }

                // map<anydata>[]
                if (elementTag == TypeTags.MAP_TAG) {
                    return parseTableToMaps(table, sheet, config, (MapType) resolvedElementType);
                }
            }

            // Default: string[][]
            return parseTableToStringArray(table, sheet, config);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing table '" + table.getName() + "': " + e.getMessage(), e);
        }
    }

    /**
     * Write data to an Excel table.
     *
     * @param table   The XSSFTable to write to
     * @param sheet   The sheet containing the table
     * @param data    Data to write
     * @param options Write options
     * @return null on success, BError on failure
     */
    public static Object writeToTable(XSSFTable table, XSSFSheet sheet, BArray data,
                                       BMap<BString, Object> options) {
        try {
            AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
            int firstRow = area.getFirstCell().getRow();
            int lastRow = area.getLastCell().getRow();
            int firstCol = area.getFirstCell().getCol();
            int lastCol = area.getLastCell().getCol();

            // Data starts after header row
            int dataFirstRow = firstRow + 1;
            int totalsRowCount = table.getTotalsRowCount();
            boolean hasTotals = totalsRowCount > 0;

            long dataSize = data.getLength();
            int currentDataRows = lastRow - firstRow - totalsRowCount;

            // Check if we need to expand the table
            if (dataSize > currentDataRows) {
                // Calculate new last row (keeping totals if present)
                int newDataLastRow = dataFirstRow + (int) dataSize - 1;
                int newLastRow = hasTotals ? newDataLastRow + 1 : newDataLastRow;

                // Resize the table
                CellReference topLeft = new CellReference(firstRow, firstCol);
                CellReference bottomRight = new CellReference(newLastRow, lastCol);
                AreaReference newArea = new AreaReference(topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);
                CTTable ctTable = table.getCTTable();
                ctTable.setRef(newArea.formatAsString());

                // Update table references
                table.updateHeaders();
            }

            // Write the data
            for (int i = 0; i < dataSize; i++) {
                int rowIdx = dataFirstRow + i;
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    row = sheet.createRow(rowIdx);
                }

                Object rowData = data.get(i);
                writeRowData(row, firstCol, rowData);
            }

            return null;

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to table '" + table.getName() + "': " + e.getMessage(), e);
        }
    }

    // === Private Helper Methods ===

    /**
     * Parse table data to string[][].
     */
    private static BArray parseTableToStringArray(XSSFTable table, XSSFSheet sheet, XlsxConfig config) {
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
     * Parse table data to record[].
     */
    private static Object parseTableToRecords(XSSFTable table, XSSFSheet sheet, XlsxConfig config,
                                               RecordType recordType) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();

        // Data starts after header row
        int dataFirstRow = firstRow + 1;
        int dataLastRow = lastRow - table.getTotalsRowCount();

        // Build header map from table columns
        List<XSSFTableColumn> columns = table.getColumns();
        Map<String, Integer> headerMap = new HashMap<>();
        for (int i = 0; i < columns.size(); i++) {
            String colName = columns.get(i).getName();
            if (config.isCaseInsensitiveHeaders()) {
                headerMap.put(colName.toLowerCase(), firstCol + i);
            } else {
                headerMap.put(colName, firstCol + i);
            }
        }

        // Get field mappings
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

        // Validate absent fields
        RecordParsingUtils.validateAbsentFields(absentFields, sheet.getSheetName(), config);

        // Parse rows
        ArrayType arrayType = TypeCreator.createArrayType(recordType);
        List<BMap<BString, Object>> records = new ArrayList<>();

        Integer rowCountLimit = config.getRowCount();
        int parsedCount = 0;

        for (int rowIdx = dataFirstRow; rowIdx <= dataLastRow; rowIdx++) {
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = sheet.getRow(rowIdx);
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
                    String cellAddress = RecordParsingUtils.getCellAddress(colIdx, rowIdx);
                    throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                            e.getMessage(), cellAddress, rowIdx, colIdx));
                }

                if (value != null) {
                    record.put(StringUtils.fromString(mapping.fieldName), value);
                } else {
                    if (config.isNilAsOptionalField() && mapping.isOptional) {
                        continue;
                    }
                    if (RecordParsingUtils.isNilableType(mapping.type) || mapping.isOptional) {
                        record.put(StringUtils.fromString(mapping.fieldName), null);
                    } else {
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
                        String cellAddress = RecordParsingUtils.getCellAddress(colIdx, rowIdx);
                        throw new BallerinaErrorException(DiagnosticLog.typeConversionError(
                                e.getMessage(), cellAddress, rowIdx, colIdx));
                    }

                    if (value != null) {
                        record.put(StringUtils.fromString(headerName), value);
                    } else if (!config.isNilAsOptionalField()) {
                        record.put(StringUtils.fromString(headerName), null);
                    }
                }
            }

            // Handle absent fields
            RecordParsingUtils.handleAbsentFields(record, absentFields, config);

            records.add(record);
            parsedCount++;
        }

        BArray result = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> record : records) {
            result.append(record);
        }

        return result;
    }

    /**
     * Parse table data to map<anydata>[].
     */
    private static Object parseTableToMaps(XSSFTable table, XSSFSheet sheet, XlsxConfig config, MapType mapType) {
        AreaReference area = new AreaReference(table.getArea().formatAsString(), SpreadsheetVersion.EXCEL2007);
        int firstRow = area.getFirstCell().getRow();
        int lastRow = area.getLastCell().getRow();
        int firstCol = area.getFirstCell().getCol();

        // Data starts after header row
        int dataFirstRow = firstRow + 1;
        int dataLastRow = lastRow - table.getTotalsRowCount();

        // Get column names from table
        List<XSSFTableColumn> columns = table.getColumns();

        ArrayType arrayType = TypeCreator.createArrayType(mapType);
        List<BMap<BString, Object>> maps = new ArrayList<>();

        Integer rowCountLimit = config.getRowCount();
        int parsedCount = 0;

        for (int rowIdx = dataFirstRow; rowIdx <= dataLastRow; rowIdx++) {
            if (rowCountLimit != null && parsedCount >= rowCountLimit) {
                break;
            }

            Row row = sheet.getRow(rowIdx);
            BMap<BString, Object> map = ValueCreator.createMapValue(mapType);

            for (int i = 0; i < columns.size(); i++) {
                String columnName = columns.get(i).getName();
                int colIdx = firstCol + i;

                Cell cell = row != null ? row.getCell(colIdx) : null;
                Object value = convertCellToAnydata(cell, config);
                map.put(StringUtils.fromString(columnName), value);
            }

            maps.add(map);
            parsedCount++;
        }

        BArray result = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> map : maps) {
            result.append(map);
        }

        return result;
    }

    /**
     * Write row data to a POI Row.
     */
    private static void writeRowData(Row row, int startCol, Object rowData) {
        if (rowData instanceof BMap) {
            @SuppressWarnings("unchecked")
            BMap<BString, Object> record = (BMap<BString, Object>) rowData;
            BString[] keys = record.getKeys();
            for (int i = 0; i < keys.length; i++) {
                Cell cell = row.createCell(startCol + i);
                Object value = record.get(keys[i]);
                CellConverter.setCellValue(cell, value);
            }
        } else if (rowData instanceof BArray) {
            BArray array = (BArray) rowData;
            for (int i = 0; i < array.getLength(); i++) {
                Cell cell = row.createCell(startCol + i);
                Object value = array.get(i);
                CellConverter.setCellValue(cell, value);
            }
        }
    }

    /**
     * Convert a cell to anydata (preserving type information).
     * For map<anydata>, we want to preserve actual types (numbers, booleans, strings).
     */
    private static Object convertCellToAnydata(Cell cell, XlsxConfig config) {
        if (cell == null) {
            return null;
        }

        org.apache.poi.ss.usermodel.CellType cellType = cell.getCellType();

        // Handle formula cells
        if (cellType == org.apache.poi.ss.usermodel.CellType.FORMULA) {
            if (config != null && config.isFormulaModeText()) {
                return StringUtils.fromString("=" + cell.getCellFormula());
            }
            cellType = cell.getCachedFormulaResultType();
        }

        switch (cellType) {
            case STRING:
                return StringUtils.fromString(cell.getStringCellValue());
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    // Return date as ISO format string for anydata
                    java.util.Date date = cell.getDateCellValue();
                    java.time.LocalDate localDate = date.toInstant()
                            .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                    return StringUtils.fromString(localDate.toString());
                }
                double numValue = cell.getNumericCellValue();
                // Return as long if it's a whole number, otherwise as decimal
                if (numValue == Math.floor(numValue) && !Double.isInfinite(numValue)) {
                    return (long) numValue;
                }
                return ValueCreator.createDecimalValue(java.math.BigDecimal.valueOf(numValue));
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
                return null;
            default:
                return StringUtils.fromString(cell.toString());
        }
    }
}
