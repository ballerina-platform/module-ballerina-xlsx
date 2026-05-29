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

import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;

import java.lang.reflect.Method;

/**
 * Utility class for constraint validation.
 */
public final class ConstraintUtils {

    // Thread-safe eager initialization of constraint module availability
    private static final boolean CONSTRAINT_MODULE_AVAILABLE;

    static {
        boolean available;
        try {
            Class.forName("io.ballerina.stdlib.constraint.Constraints");
            available = true;
        } catch (ClassNotFoundException e) {
            available = false;
        }
        CONSTRAINT_MODULE_AVAILABLE = available;
    }

    private ConstraintUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Validate a record against its type constraints.
     *
     * @param record     The record to validate
     * @param recordType The record type containing constraints
     * @return The validated record if successful, or BError if validation fails
     */
    public static Object validate(BMap<BString, Object> record, RecordType recordType) {
        // Check if constraint module is available
        if (!isConstraintModuleAvailable()) {
            // Constraint module not available, skip validation
            return record;
        }

        try {
            // Use reflection to call constraint:validate() function
            // This avoids compile-time dependency on the constraint module
            Class<?> constraintClass = Class.forName("io.ballerina.stdlib.constraint.Constraints");
            Method validateMethod = constraintClass.getMethod("validate", Object.class, BTypedesc.class);

            // Create a typedesc for the record type
            BTypedesc typedesc = ValueCreator.createTypedescValue(recordType);

            // Call the validate function
            Object result = validateMethod.invoke(null, record, typedesc);

            return result;
        } catch (Exception e) {
            // If validation call fails, return the error
            if (e.getCause() instanceof BError) {
                return e.getCause();
            }
            // Wrap other exceptions as constraint validation error
            return DiagnosticLog.constraintValidationError(
                    "Constraint validation failed: " + e.getMessage(), null, null);
        }
    }

    /**
     * Check if the constraint module is available.
     *
     * @return true if constraint module is available
     */
    private static boolean isConstraintModuleAvailable() {
        return CONSTRAINT_MODULE_AVAILABLE;
    }
}
