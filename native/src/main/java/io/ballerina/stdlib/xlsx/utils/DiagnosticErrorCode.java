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

package io.ballerina.stdlib.xlsx.utils;

/**
 * Diagnostic error codes for XLSX module.
 */
public enum DiagnosticErrorCode {

    // Parse errors (1xx)
    INVALID_XLSX_FORMAT("XLSX_ERROR_101", "invalid.xlsx.format"),
    XLSX_PARSE_FAILED("XLSX_ERROR_102", "xlsx.parse.failed"),
    STREAM_READ_FAILED("XLSX_ERROR_103", "stream.read.failed"),

    // Sheet errors (2xx)
    SHEET_NOT_FOUND("XLSX_ERROR_201", "sheet.not.found"),
    SHEET_INDEX_OUT_OF_BOUNDS("XLSX_ERROR_202", "sheet.index.out.of.bounds"),
    SHEET_ALREADY_EXISTS("XLSX_ERROR_203", "sheet.already.exists"),

    // Type conversion errors (3xx)
    TYPE_CONVERSION_FAILED("XLSX_ERROR_301", "type.conversion.failed"),
    UNSUPPORTED_TYPE("XLSX_ERROR_302", "unsupported.type"),
    CELL_TYPE_MISMATCH("XLSX_ERROR_303", "cell.type.mismatch"),

    // Resource errors (4xx)
    FILE_TOO_LARGE("XLSX_ERROR_401", "file.too.large"),
    TOO_MANY_ROWS("XLSX_ERROR_402", "too.many.rows"),
    RESOURCE_EXHAUSTED("XLSX_ERROR_403", "resource.exhausted"),

    // Write errors (5xx)
    WRITE_FAILED("XLSX_ERROR_501", "write.failed"),
    SERIALIZATION_FAILED("XLSX_ERROR_502", "serialization.failed"),

    // Fail-safe errors (6xx)
    FAILED_FILE_IO_OPERATION("XLSX_ERROR_601", "failed.file.io.operation"),
    INVALID_XLSX_DATA_FORMAT("XLSX_ERROR_602", "invalid.xlsx.data.format"),
    NO_FIELD_FOR_HEADER("XLSX_ERROR_603", "no.field.for.header"),
    HEADER_CANNOT_BE_EMPTY("XLSX_ERROR_604", "header.cannot.be.empty"),

    // Table errors (7xx)
    TABLE_NOT_FOUND("XLSX_ERROR_701", "table.not.found"),
    TABLE_OVERLAP("XLSX_ERROR_702", "table.overlap"),
    INVALID_TABLE_RANGE("XLSX_ERROR_703", "invalid.table.range"),
    TABLE_ALREADY_EXISTS("XLSX_ERROR_704", "table.already.exists");

    private final String errorCode;
    private final String messageKey;

    DiagnosticErrorCode(String errorCode, String messageKey) {
        this.errorCode = errorCode;
        this.messageKey = messageKey;
    }

    public String getErrorCode() {
        return errorCode;
    }

    public String getMessageKey() {
        return messageKey;
    }
}
