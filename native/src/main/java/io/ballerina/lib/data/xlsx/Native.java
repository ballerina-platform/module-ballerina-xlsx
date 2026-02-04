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
import io.ballerina.runtime.api.values.BStream;
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
    public static Object parse(Environment env, BString filePath, Object sheet,
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
    public static Object write(BArray data, BString filePath, BMap<BString, Object> options) {
        return XlsxWriter.writeToFile(filePath.getValue(), data, options);
    }

    /**
     * Parse XLSX from a byte stream into a Ballerina array.
     * Note: Deferred to v2 - XLSX format requires SharedStringsTable, making true streaming complex.
     *
     * @param env        Ballerina environment
     * @param xlsxStream The XLSX content as a byte stream
     * @param sheet      Sheet to read (string name or int index)
     * @param options    Parsing options
     * @param typedesc   Target type descriptor
     * @return Parsed data as BArray or error
     */
    public static Object parseAsStream(Environment env, BStream xlsxStream, Object sheet,
                                       BMap<BString, Object> options, BTypedesc typedesc) {
        // Deferred to v2: XLSX format requires SharedStringsTable to be loaded first
        return DiagnosticLog.error("parseAsStream is deferred to v2. Use parse() instead.");
    }
}
