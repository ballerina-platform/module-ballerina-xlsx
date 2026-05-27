/*
 * Copyright (c) 2025, WSO2 LLC. (https://www.wso2.com).
 *
 * WSO2 LLC. licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package io.ballerina.stdlib.xlsx.xlsx;

import io.ballerina.runtime.api.values.BError;

/**
 * Custom exception that wraps a Ballerina BError to preserve type information.
 * <p>
 * This allows throwing errors through Java code while maintaining the full
 * BError structure including error type (ParseError, TypeConversionError, etc.)
 * and structured details (sheetName, cellAddress, rowNumber, columnNumber).
 * </p>
 *
 * @since 1.0.0
 */
public class BallerinaErrorException extends RuntimeException {

    private final BError bError;

    /**
     * Create a new BallerinaErrorException wrapping a BError.
     *
     * @param bError The Ballerina error to wrap
     */
    public BallerinaErrorException(BError bError) {
        super(bError.getMessage());
        this.bError = bError;
    }

    /**
     * Get the wrapped Ballerina error.
     *
     * @return The original BError object with type and details intact
     */
    public BError getBError() {
        return bError;
    }
}
