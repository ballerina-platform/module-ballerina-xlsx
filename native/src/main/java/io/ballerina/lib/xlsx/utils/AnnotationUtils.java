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

import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

/**
 * Utility class for handling Ballerina annotations in XLSX operations.
 */
public final class AnnotationUtils {

    private AnnotationUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Get the header name for a field, checking for @xlsx:Name annotation.
     *
     * @param recordType The record type containing annotations
     * @param fieldName  The field name
     * @return The header name from annotation, or the field name if no annotation
     */
    public static String getHeaderName(RecordType recordType, String fieldName) {
        BMap<BString, Object> annotations = recordType.getAnnotations();
        if (annotations == null || annotations.isEmpty()) {
            return fieldName;
        }

        // Annotation structure (nested):
        // - Outer key: "$field$.{fieldName}"
        // - Inner key: "{org}/{module}:{version}:{annotationName}" e.g., "ballerina/xlsx:0:Name"
        // - Inner value: Map with "value" key
        String fieldKey = "$field$." + fieldName;
        Object fieldAnnotations = annotations.get(StringUtils.fromString(fieldKey));
        if (!(fieldAnnotations instanceof BMap)) {
            return fieldName;
        }

        @SuppressWarnings("unchecked")
        BMap<BString, Object> fieldAnnotMap = (BMap<BString, Object>) fieldAnnotations;

        // Look for any annotation ending with ":Name" that references xlsx module
        for (BString annotKey : fieldAnnotMap.getKeys()) {
            String annotKeyStr = annotKey.getValue();
            if (!annotKeyStr.endsWith(":Name") || !annotKeyStr.contains("xlsx")) {
                continue;
            }
            Object annotValue = fieldAnnotMap.get(annotKey);
            if (!(annotValue instanceof BMap)) {
                continue;
            }
            @SuppressWarnings("unchecked")
            BMap<BString, Object> annotMap = (BMap<BString, Object>) annotValue;
            Object value = annotMap.get(StringUtils.fromString("value"));
            if (value != null) {
                // Trim spurious whitespace so " Name " in the annotation matches a
                // sheet header of "Name". The sheet-side trim is already done in
                // RecordParsingUtils.buildHeaderMap.
                return value.toString().trim();
            }
        }

        return fieldName;
    }
}
