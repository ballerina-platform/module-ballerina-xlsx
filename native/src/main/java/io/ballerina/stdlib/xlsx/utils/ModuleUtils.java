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

import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.Module;

import java.util.HashMap;
import java.util.Map;

/**
 * Utility class for module-related operations.
 */
public final class ModuleUtils {

    private static Module module;

    private ModuleUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Set the module from Ballerina environment.
     * Called during module initialization.
     *
     * @param env Ballerina environment
     */
    public static void setModule(Environment env) {
        module = env.getCurrentModule();
    }

    /**
     * Get the current module.
     *
     * @return Module instance
     */
    public static Module getModule() {
        return module;
    }

    /**
     * Get properties map for StrandMetadata.
     * Used when calling Ballerina functions from Java.
     *
     * @param functionName The name of the function being called
     * @return Map of properties
     */
    public static Map<String, Object> getProperties(String functionName) {
        Map<String, Object> properties = new HashMap<>();
        properties.put("functionName", functionName);
        return properties;
    }
}
