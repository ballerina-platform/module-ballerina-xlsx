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
import io.ballerina.stdlib.xlsx.utils.DiagnosticLog;
import io.ballerina.stdlib.xlsx.utils.RecordParsingUtils;
import io.ballerina.stdlib.xlsx.utils.UsedRangeDetector;
import io.ballerina.stdlib.xlsx.utils.XlsxConfig;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Core XLSX parsing logic using Apache POI.
 */
public final class XlsxParser {

    private XlsxParser() {
        // Private constructor to prevent instantiation
    }

    /**
     * Parse XLSX bytes to Ballerina value based on target type.
     *
     * @param env        Ballerina environment (for fail-safe logging)
     * @param data       XLSX file bytes
     * @param sheet      Sheet to read (BString for name, Long for index)
     * @param options    Parse options
     * @param targetType Target Ballerina type descriptor
     * @return Parsed value (array of records or array of arrays)
     */
    public static Object parseBytes(Environment env, byte[] data, Object sheet, BMap<BString, Object> options,
                                    BTypedesc targetType) {
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try (ByteArrayInputStream bis = new ByteArrayInputStream(data);
             Workbook workbook = WorkbookFactory.create(bis)) {

            Sheet selectedSheet = selectSheet(workbook, sheet);
            return parseSheet(env, selectedSheet, config, targetType);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (IOException e) {
            return DiagnosticLog.parseError("Failed to parse XLSX: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing XLSX file: " + e.getMessage(), e);
        }
    }

    /**
     * Parse XLSX from file path (memory efficient - avoids double loading).
     *
     * @param env        Ballerina environment (for fail-safe logging)
     * @param filePath   Path to the XLSX file
     * @param sheet      Sheet to read (BString for name, Long for index)
     * @param options    Parse options
     * @param targetType Target Ballerina type descriptor
     * @return Parsed value (array of records or array of arrays)
     */
    public static Object parseFromFile(Environment env, Path filePath, Object sheet,
                                       BMap<BString, Object> options, BTypedesc targetType) {
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        // Pass File directly to WorkbookFactory for better efficiency:
        // POI can use ZipFile with random access via FileChannel instead of
        // sequential ZipInputStream, avoiding intermediate buffering.
        try (Workbook workbook = WorkbookFactory.create(filePath.toFile())) {

            Sheet selectedSheet = selectSheet(workbook, sheet);
            return parseSheet(env, selectedSheet, config, targetType);

        } catch (BallerinaErrorException e) {
            return e.getBError();
        } catch (IOException e) {
            return DiagnosticLog.parseError("Failed to parse XLSX: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing XLSX file: " + e.getMessage(), e);
        }
    }

    /**
     * Select sheet from workbook based on sheet identifier.
     *
     * @param workbook The workbook to select from
     * @param sheet    Sheet identifier (BString for name, Long for index)
     * @return Selected sheet
     * @throws BallerinaErrorException if sheet is not found
     */
    private static Sheet selectSheet(Workbook workbook, Object sheet) {
        Sheet selectedSheet;

        if (sheet instanceof BString) {
            String sheetName = ((BString) sheet).getValue();
            selectedSheet = workbook.getSheet(sheetName);
            if (selectedSheet == null) {
                throw new BallerinaErrorException(DiagnosticLog.sheetNotFoundError(sheetName));
            }
        } else if (sheet instanceof Long) {
            int index = ((Long) sheet).intValue();
            if (index < 0 || index >= workbook.getNumberOfSheets()) {
                throw new BallerinaErrorException(
                        DiagnosticLog.sheetNotFoundError(index, workbook.getNumberOfSheets() - 1));
            }
            selectedSheet = workbook.getSheetAt(index);
        } else {
            // Default: first sheet
            selectedSheet = workbook.getSheetAt(0);
        }

        return selectedSheet;
    }

    /**
     * Parse a sheet to the target type.
     */
    private static Object parseSheet(Environment env, Sheet sheet, XlsxConfig config, BTypedesc targetType) {
        // Under the typedesc<Row> signature the describing type IS the row element type
        // (the function returns `t[]`, so Ballerina infers t = element). Dispatch directly
        // on the row shape — no need to unwrap an outer array first.
        Type describingType = TypeUtils.getReferredType(targetType.getDescribingType());
        int typeTag = describingType.getTag();

        // string[] - raw string row → string[][]
        if (typeTag == TypeTags.ARRAY_TAG) {
            ArrayType arrayType = (ArrayType) describingType;
            Type innerElementType = TypeUtils.getReferredType(arrayType.getElementType());
            if (innerElementType.getTag() == TypeTags.STRING_TAG) {
                return parseToStringArray(sheet, config);
            }
        }

        // record{} - array of records
        if (typeTag == TypeTags.RECORD_TYPE_TAG) {
            return parseToRecordArray(env, sheet, config, (RecordType) describingType);
        }

        // map<anydata> - array of maps
        if (typeTag == TypeTags.MAP_TAG) {
            return parseToMapArray(env, sheet, config, (MapType) describingType);
        }

        // Row (the public union) — the caller asked for the abstract row shape without
        // narrowing. Return the most general usable representation: string[][].
        if (typeTag == TypeTags.UNION_TAG) {
            return parseToStringArray(sheet, config);
        }

        // Default: parse as string array
        return parseToStringArray(sheet, config);
    }

    /**
     * Parse sheet to string[][].
     */
    private static BArray parseToStringArray(Sheet sheet, XlsxConfig config) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        if (usedRange == null) {
            // Empty sheet - return empty array
            ArrayType stringArrayType = TypeCreator.createArrayType(
                    TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING));
            return ValueCreator.createArrayValue(stringArrayType);
        }

        int startRow = config.hasExplicitDataStartRowIndex() ?
                config.getDataStartRowIndex() : usedRange.getFirstRow();
        int endRow = usedRange.getLastRow();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        // Apply rowCount limit if set
        Integer rowCountLimit = config.getRowCount();

        List<BArray> rows = new ArrayList<>();
        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
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

        ArrayType resultType = TypeCreator.createArrayType(stringType);
        BArray result = ValueCreator.createArrayValue(resultType);
        for (BArray row : rows) {
            result.append(row);
        }

        return result;
    }

    /**
     * Parse sheet to record[].
     * Returns BArray on success, or BError on failure.
     */
    private static Object parseToRecordArray(Environment env, Sheet sheet, XlsxConfig config, RecordType recordType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, recordType, env, isOverwritten);

        return RecordParsingUtils.parseRowsToRecords(context);
    }

    /**
     * Parse sheet to map<anydata>[].
     * Returns BArray on success, or BError on failure.
     */
    private static Object parseToMapArray(Environment env, Sheet sheet, XlsxConfig config, MapType mapType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);
        AtomicBoolean isOverwritten = new AtomicBoolean(false);

        RecordParsingUtils.ParseContext context = new RecordParsingUtils.ParseContext(
                sheet, usedRange, config, mapType, env, isOverwritten);

        return RecordParsingUtils.parseRowsToMaps(context);
    }
}
