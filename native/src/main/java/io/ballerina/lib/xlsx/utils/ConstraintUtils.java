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

import io.ballerina.lib.xlsx.xlsx.BallerinaErrorException;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;

import java.lang.reflect.Method;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Utility class for constraint validation.
 */
public final class ConstraintUtils {

    // Fully-qualified name of the optional constraint module's entry point and its validate method,
    // resolved reflectively so the module stays an optional (compile-time-absent) dependency.
    private static final String CONSTRAINT_CLASS_NAME = "io.ballerina.stdlib.constraint.Constraints";
    private static final String VALIDATE_METHOD_NAME = "validate";

    // Thread-safe eager initialization of constraint module availability and the resolved
    // validate(Object, BTypedesc) method, so the reflective lookup happens once, not per row.
    private static final boolean CONSTRAINT_MODULE_AVAILABLE;
    private static final Method VALIDATE_METHOD;

    // Typedescs are immutable; cache one per record type rather than rebuilding it per row.
    private static final Map<RecordType, BTypedesc> TYPEDESC_CACHE = new ConcurrentHashMap<>();

    static {
        boolean available;
        Method method = null;
        try {
            Class<?> constraintClass = Class.forName(CONSTRAINT_CLASS_NAME);
            method = constraintClass.getMethod(VALIDATE_METHOD_NAME, Object.class, BTypedesc.class);
            available = true;
        } catch (ClassNotFoundException | NoSuchMethodException e) {
            available = false;
        }
        CONSTRAINT_MODULE_AVAILABLE = available;
        VALIDATE_METHOD = method;
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
        // Constraint module not on the classpath → nothing to validate against.
        if (!isConstraintModuleAvailable()) {
            return record;
        }

        try {
            // The validate(Object, BTypedesc) method and the per-type typedesc are both cached, so
            // the hot path is a single reflective invoke per row.
            BTypedesc typedesc = TYPEDESC_CACHE.computeIfAbsent(recordType,
                    rt -> ValueCreator.createTypedescValue(rt));
            return VALIDATE_METHOD.invoke(null, record, typedesc);
        } catch (Exception e) {
            // `e` is the InvocationTargetException from the reflective invoke; its own message is not
            // meaningful. The real failure is its cause — a BError carrying the constraint message — so
            // return that (the caller records it and skips the row under fail-safe).
            if (e.getCause() instanceof BError) {
                return e.getCause();
            }
            // Anything else (a reflection/wiring failure, not a per-row data problem) must fail fast
            // rather than be masked as a skippable validation error.
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Constraint validation could not be performed: " + e.getMessage(), e));
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
