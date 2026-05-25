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

package io.ballerina.lib.data.xlsx;

import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.xlsx.XlsxParser;
import io.ballerina.lib.data.xlsx.xlsx.XlsxWriter;
import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Native entry point for Ballerina XLSX module.
 * This class provides the bridge between Ballerina and Java for Excel operations.
 */
public final class Native {

    private Native() {
        // Private constructor to prevent instantiation
    }

    /**
     * Parse XLSX file from a file path into a Ballerina array.
     * This is the PRIMARY API for parsing XLSX files.
     *
     * @param env       Ballerina environment (for fail-safe logging)
     * @param filePath  Path to the XLSX file
     * @param sheet     Sheet to read (string name or int index)
     * @param options   Parsing options
     * @param typedesc  Target type descriptor
     * @return Parsed data as BArray or error
     */
    public static Object parseSheet(Environment env, BString filePath, Object sheet,
                                    BMap<BString, Object> options, BTypedesc typedesc) {
        Path path = Paths.get(filePath.getValue());

        // Check if file exists before parsing
        if (!Files.exists(path)) {
            return DiagnosticLog.fileNotFoundError("Failed to read file: " + filePath.getValue(),
                    new java.io.FileNotFoundException(filePath.getValue()));
        }

        return XlsxParser.parseFromFile(env, path, sheet, options, typedesc);
    }

    /**
     * Write Ballerina data directly to an XLSX file.
     * This is the PRIMARY API for writing XLSX files.
     *
     * @param data     Data to write (record[] or string[][])
     * @param filePath Path to the output file
     * @param options  Write options
     * @return null on success, error on failure
     */
    public static Object writeSheet(BArray data, BString filePath, BMap<BString, Object> options) {
        return XlsxWriter.writeToFile(filePath.getValue(), data, options);
    }

    // ==========================================================================
    // TABLE API
    // ==========================================================================

    /**
     * Parse data from an Excel table by name.
     *
     * @param env       Ballerina environment
     * @param filePath  Path to the XLSX file
     * @param tableName Name of the table to parse
     * @param options   Parsing options
     * @param typedesc  Target type descriptor
     * @return Parsed data as BArray or error
     */
    public static Object parseTable(Environment env, BString filePath, BString tableName,
                                    BMap<BString, Object> options, BTypedesc typedesc) {
        Path path = Paths.get(filePath.getValue());

        // Check if file exists before parsing
        if (!Files.exists(path)) {
            return DiagnosticLog.fileNotFoundError("Failed to read file: " + filePath.getValue(),
                    new java.io.FileNotFoundException(filePath.getValue()));
        }

        try {
            // Open workbook, find table, parse data
            org.apache.poi.ss.usermodel.Workbook workbook =
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(path.toFile());

            if (!(workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook)) {
                workbook.close();
                return DiagnosticLog.error("Tables are only supported in XLSX format");
            }

            org.apache.poi.xssf.usermodel.XSSFWorkbook xssfWorkbook =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook;
            String name = tableName.getValue();

            // Search all sheets for the table
            for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet =
                        (org.apache.poi.xssf.usermodel.XSSFSheet) xssfWorkbook.getSheetAt(i);
                for (org.apache.poi.xssf.usermodel.XSSFTable table : sheet.getTables()) {
                    if (name.equals(table.getName()) || name.equals(table.getDisplayName())) {
                        // Found the table - parse its data
                        Object result = io.ballerina.lib.data.xlsx.xlsx.TableParser.parseTable(
                                env, table, sheet, options, typedesc);
                        workbook.close();
                        return result;
                    }
                }
            }

            workbook.close();
            return DiagnosticLog.tableNotFoundError(name);

        } catch (java.io.IOException e) {
            return DiagnosticLog.error("Failed to parse XLSX file: " + e.getMessage(), e);
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing table: " + e.getMessage(), e);
        }
    }

    /**
     * Write data to an existing Excel table.
     *
     * @param data      Data to write
     * @param filePath  Path to the XLSX file
     * @param tableName Name of the table to write to
     * @param options   Write options
     * @return null on success, error on failure
     */
    public static Object writeTable(BArray data, BString filePath, BString tableName,
                                    BMap<BString, Object> options) {
        Path path = Paths.get(filePath.getValue());

        // Check if file exists
        if (!Files.exists(path)) {
            return DiagnosticLog.fileNotFoundError("Failed to read file: " + filePath.getValue(),
                    new java.io.FileNotFoundException(filePath.getValue()));
        }

        try {
            // Open workbook for modification
            java.io.FileInputStream fis = new java.io.FileInputStream(path.toFile());
            org.apache.poi.ss.usermodel.Workbook workbook =
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(fis);
            fis.close();

            if (!(workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook)) {
                workbook.close();
                return DiagnosticLog.error("Tables are only supported in XLSX format");
            }

            org.apache.poi.xssf.usermodel.XSSFWorkbook xssfWorkbook =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook;
            String name = tableName.getValue();

            // Search all sheets for the table
            for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet =
                        (org.apache.poi.xssf.usermodel.XSSFSheet) xssfWorkbook.getSheetAt(i);
                for (org.apache.poi.xssf.usermodel.XSSFTable table : sheet.getTables()) {
                    if (name.equals(table.getName()) || name.equals(table.getDisplayName())) {
                        // Found the table - write data to it
                        Object result = io.ballerina.lib.data.xlsx.xlsx.TableParser.writeToTable(
                                table, sheet, data, options);
                        if (result != null) {
                            workbook.close();
                            return result;  // Error
                        }

                        // Save the workbook
                        java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile());
                        workbook.write(fos);
                        fos.close();
                        workbook.close();
                        return null;  // Success
                    }
                }
            }

            workbook.close();
            return DiagnosticLog.tableNotFoundError(name);

        } catch (java.io.IOException e) {
            return DiagnosticLog.error("Failed to write to XLSX file: " + e.getMessage(), e);
        } catch (Exception e) {
            return DiagnosticLog.error("Error writing to table: " + e.getMessage(), e);
        }
    }
}
