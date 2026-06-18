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

package io.ballerina.lib.xlsx.utils;

import io.ballerina.lib.xlsx.xlsx.TypeConversionException;
import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.concurrent.StrandMetadata;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Utility class for fail-safe error handling and logging in XLSX operations.
 */
public final class FailSafeUtils {

    // Ballerina error type names that fail-safe may recover from (skip the row + log). Everything
    // else — unexpected POI/runtime faults, IO errors, structural Ballerina errors — fails fast.
    private static final String TYPE_CONVERSION_ERROR_TYPE = "TypeConversionError";
    private static final String CONSTRAINT_VALIDATION_ERROR_TYPE = "ConstraintValidationError";

    // Fail-safe options field names
    private static final BString FILE_OUTPUT_MODE = StringUtils.fromString("fileOutputMode");
    private static final BString FILE_PATH = StringUtils.fromString("filePath");
    private static final BString CONTENT_TYPE = StringUtils.fromString("contentType");
    private static final BString FILE_WRITE_OPTION = StringUtils.fromString("fileWriteOption");
    private static final BString OFFENDING_ROW = StringUtils.fromString("offendingRow");

    // Fail-safe content type values
    private static final String CONTENT_TYPE_METADATA = "METADATA";
    private static final String CONTENT_TYPE_RAW = "RAW";

    // Fail-safe file write option value
    private static final String FILE_WRITE_OVERWRITE = "OVERWRITE";

    // Name of the Ballerina function used to emit fail-safe warnings
    private static final String PRINT_FAILSAFE_WARNING = "printFailSafeWarning";

    // Fail-safe warning and error messages
    private static final String XLSX_PARSE_WARNING = "XLSX parse warning at row %d, column %d: %s";
    private static final String FILE_IO_ERROR = "Failed to create log file at: %s. Caused by: %s";
    private static final String FILE_OVERWRITE_ERROR = "Failed to overwrite log file at: %s. Caused by: %s";
    private static final String FILE_WRITE_ERROR = "Failed to write log file at: %s. Caused by: %s";

    private FailSafeUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Whether an exception is recoverable under fail-safe (skip the offending row + log) rather
     * than failing the whole parse.
     *
     * <p>This is a strict allowlist: only a per-row data fault — a cell that cannot be converted
     * to its target type ({@link TypeConversionException} or a {@code TypeConversionError}), or a
     * row that fails constraint validation ({@code ConstraintValidationError}) — is recoverable.
     * Anything else (unexpected POI/runtime exceptions, IO errors, and structural Ballerina errors
     * such as a missing sheet or an empty header row) fails fast, so internal faults are never
     * silently masked as skipped rows.</p>
     *
     * @param exception The exception to classify
     * @return true if the row may be skipped under fail-safe, false if it must fail immediately
     */
    public static boolean isAllowedFailSafe(Exception exception) {
        if (exception instanceof TypeConversionException) {
            return true;
        }
        if (exception instanceof BError bError) {
            String typeName = bError.getType().getName();
            return TYPE_CONVERSION_ERROR_TYPE.equals(typeName)
                    || CONSTRAINT_VALIDATION_ERROR_TYPE.equals(typeName);
        }
        return false;
    }

    /**
     * Handle fail-safe logging to console and/or file.
     *
     * @param environment              The Ballerina environment
     * @param failSafe                 The fail-safe configuration map
     * @param exception                The exception that occurred
     * @param offendingRow             The row data that caused the error (as JSON array string)
     * @param rowIndex                 Row index (0-based)
     * @param columnIndex              Column index (0-based)
     * @param isOverwritten            Atomic flag tracking if file has been overwritten
     * @param enableConsoleLogs        Whether console logging is enabled
     * @param includeSourceDataInConsole Whether to include source data in console logs
     */
    public static void handleFailSafeLogging(Environment environment, BMap<?, ?> failSafe,
                                             Exception exception, String offendingRow,
                                             int rowIndex, int columnIndex,
                                             AtomicBoolean isOverwritten,
                                             boolean enableConsoleLogs,
                                             boolean includeSourceDataInConsole) {
        // Console logging
        if (enableConsoleLogs && environment != null) {
            processConsoleLogs(environment, exception, includeSourceDataInConsole, offendingRow,
                    rowIndex, columnIndex);
        }

        // File logging
        BMap<?, ?> fileOutputMode = failSafe.getMapValue(FILE_OUTPUT_MODE);
        if (fileOutputMode != null) {
            processErrorLogsInFiles(exception, fileOutputMode, offendingRow, rowIndex, columnIndex, isOverwritten);
        }
    }

    /**
     * Process console logging.
     *
     * @param environment       The Ballerina environment
     * @param exception         The exception
     * @param includeSourceData Whether to include source data
     * @param offendingRow      The offending row data
     * @param rowIndex          Row index (0-based)
     * @param columnIndex       Column index (0-based)
     */
    public static void processConsoleLogs(Environment environment, Exception exception,
                                          boolean includeSourceData, String offendingRow,
                                          int rowIndex, int columnIndex) {
        BMap<BString, Object> keyValues = ValueCreator.createMapValue();
        if (includeSourceData && offendingRow != null) {
            keyValues.put(OFFENDING_ROW, StringUtils.fromString(offendingRow.trim()));
        }
        printFailSafeWarning(environment, exception, keyValues, rowIndex, columnIndex);
    }

    /**
     * Print a fail-safe warning using Ballerina's log module via the printFailSafeWarning function.
     *
     * @param environment The Ballerina environment
     * @param exception   The exception
     * @param keyValues   Key-value pairs for structured logging
     * @param rowIndex    Row index (0-based)
     * @param columnIndex Column index (0-based)
     */
    public static void printFailSafeWarning(Environment environment, Exception exception,
                                            BMap<BString, Object> keyValues, int rowIndex, int columnIndex) {
        StrandMetadata strandMetadata = new StrandMetadata(true,
                ModuleUtils.getProperties(PRINT_FAILSAFE_WARNING));
        // Convert to 1-based indexing for display
        String warningMessage = String.format(
                XLSX_PARSE_WARNING, rowIndex + 1, columnIndex + 1, exception.getMessage());
        Object[] arguments = new Object[]{StringUtils.fromString(warningMessage), null, null, keyValues};
        environment.getRuntime().callFunction(
                ModuleUtils.getModule(), PRINT_FAILSAFE_WARNING, strandMetadata, arguments);
    }

    /**
     * Process error logging to files.
     *
     * @param exception     The exception
     * @param outputMode    The file output mode configuration
     * @param offendingRow  The offending row data
     * @param rowIndex      Row index (0-based)
     * @param columnIndex   Column index (0-based)
     * @param isOverwritten Atomic flag tracking if file has been overwritten
     */
    public static void processErrorLogsInFiles(Exception exception, BMap<?, ?> outputMode,
                                               String offendingRow, int rowIndex, int columnIndex,
                                               AtomicBoolean isOverwritten) {
        String contentType = outputMode.getStringValue(CONTENT_TYPE).toString();
        handleFileOutputLogging(outputMode, exception, offendingRow, contentType, rowIndex, columnIndex, isOverwritten);
    }

    /**
     * Handle file output logging.
     *
     * @param logFileConfig The log file configuration
     * @param exception     The exception
     * @param rawData       The raw row data
     * @param contentType   The content type (METADATA, RAW, RAW_AND_METADATA)
     * @param rowIndex      Row index (0-based)
     * @param columnIndex   Column index (0-based)
     * @param isOverwritten Atomic flag tracking if file has been overwritten
     */
    public static void handleFileOutputLogging(BMap<?, ?> logFileConfig, Exception exception,
                                               String rawData, String contentType,
                                               int rowIndex, int columnIndex,
                                               AtomicBoolean isOverwritten) {
        BString filePathValue = logFileConfig.getStringValue(FILE_PATH);
        String filePath = filePathValue != null ? filePathValue.toString() : null;

        if (filePath == null || filePath.isEmpty()) {
            return;  // No file path specified, skip file logging
        }

        // Ensure log file exists
        handleLogFileGeneration(filePath);

        // Handle OVERWRITE mode
        String fileWriteOption = logFileConfig.getStringValue(FILE_WRITE_OPTION).toString();
        if (FILE_WRITE_OVERWRITE.equals(fileWriteOption) && !isOverwritten.get()) {
            overwriteLogFile(filePath);
            isOverwritten.set(true);
        }

        // Write log content based on content type
        if (contentType.equals(CONTENT_TYPE_RAW)) {
            // RAW mode: just write the offending row
            writeLogsToFile(filePath, rawData + System.lineSeparator());
        } else {
            // METADATA or RAW_AND_METADATA: write JSON
            boolean excludeSourceData = contentType.equals(CONTENT_TYPE_METADATA);
            String jsonLog = buildJsonLog(exception, rawData, excludeSourceData, rowIndex, columnIndex);
            writeLogsToFile(filePath, jsonLog + System.lineSeparator());
        }
    }

    /**
     * Ensure the log file and parent directories exist.
     *
     * @param filePath Path to the log file
     */
    public static void handleLogFileGeneration(String filePath) {
        try {
            Path path = Paths.get(filePath);
            Path parentDir = path.getParent();
            if (parentDir != null) {
                Files.createDirectories(parentDir);
            }
            if (Files.notExists(path)) {
                Files.createFile(path);
            }
        } catch (IOException ioException) {
            throw DiagnosticLog.error(String.format(FILE_IO_ERROR, filePath, ioException.getMessage()));
        }
    }

    /**
     * Overwrite (truncate) the log file.
     *
     * @param filePath Path to the log file
     */
    public static void overwriteLogFile(String filePath) {
        Path path = Paths.get(filePath);
        try {
            Files.write(path, new byte[0], StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
        } catch (IOException exception) {
            throw DiagnosticLog.error(String.format(FILE_OVERWRITE_ERROR, filePath, exception.getMessage()));
        }
    }

    /**
     * Write log content to file.
     *
     * @param filePath Path to the log file
     * @param content  Content to write
     */
    public static void writeLogsToFile(String filePath, String content) {
        Path path = Paths.get(filePath);
        try {
            Files.writeString(path, content.trim() + System.lineSeparator(),
                    StandardOpenOption.CREATE, StandardOpenOption.WRITE, StandardOpenOption.APPEND);
        } catch (IOException exception) {
            throw DiagnosticLog.error(String.format(FILE_WRITE_ERROR, filePath, exception.getMessage()));
        }
    }

    /**
     * Build JSON log entry.
     *
     * @param exception         The exception
     * @param sourceData        The source row data
     * @param excludeSourceData Whether to exclude source data from the log
     * @param rowIndex          Row index (0-based)
     * @param columnIndex       Column index (0-based)
     * @return JSON string for the log entry
     */
    public static String buildJsonLog(Exception exception, String sourceData, boolean excludeSourceData,
                                      int rowIndex, int columnIndex) {
        String time = Instant.now().toString();
        String message = exception.getMessage() != null ? exception.getMessage() : "Unknown error";

        StringBuilder json = new StringBuilder();
        json.append("{");
        json.append("\"time\":\"").append(escapeJson(time)).append("\",");
        json.append("\"location\":{");
        json.append("\"row\":").append(rowIndex + 1);  // 1-based for display
        json.append(",\"column\":").append(columnIndex + 1);
        json.append("},");
        if (!excludeSourceData && sourceData != null) {
            json.append("\"offendingRow\":\"").append(escapeJson(sourceData.trim())).append("\",");
        }
        json.append("\"message\":\"").append(escapeJson(message)).append("\"");
        json.append("}");

        return json.toString();
    }

    /**
     * Escape special characters for JSON string.
     * Handles all JSON-required escapes including control characters.
     *
     * @param str The string to escape
     * @return Escaped string
     */
    public static String escapeJson(String str) {
        if (str == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder(str.length());
        for (char c : str.toCharArray()) {
            switch (c) {
                case '\\':
                    sb.append("\\\\");
                    break;
                case '"':
                    sb.append("\\\"");
                    break;
                case '\n':
                    sb.append("\\n");
                    break;
                case '\r':
                    sb.append("\\r");
                    break;
                case '\t':
                    sb.append("\\t");
                    break;
                case '\b':
                    sb.append("\\b");
                    break;
                case '\f':
                    sb.append("\\f");
                    break;
                default:
                    if (c < 0x20) {
                        // Escape other control characters as Unicode
                        sb.append(String.format("\\u%04x", (int) c));
                    } else {
                        sb.append(c);
                    }
            }
        }
        return sb.toString();
    }
}
