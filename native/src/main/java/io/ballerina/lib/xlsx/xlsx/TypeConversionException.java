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

/**
 * Exception thrown when a cell value cannot be converted to the target type.
 */
public class TypeConversionException extends RuntimeException {

    private final String cellValue;
    private final String targetType;
    private final String actualType;

    public TypeConversionException(String message, String cellValue, String targetType, String actualType) {
        super(message);
        this.cellValue = cellValue;
        this.targetType = targetType;
        this.actualType = actualType;
    }

    public TypeConversionException(String message, String cellValue, String targetType, String actualType,
                                   Throwable cause) {
        super(message, cause);
        this.cellValue = cellValue;
        this.targetType = targetType;
        this.actualType = actualType;
    }

    public String getCellValue() {
        return cellValue;
    }

    public String getTargetType() {
        return targetType;
    }

    public String getActualType() {
        return actualType;
    }
}
