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

package io.ballerina.lib.data.xlsx.utils;

import io.ballerina.runtime.api.types.Field;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.types.UnionType;
import io.ballerina.runtime.api.utils.TypeUtils;

import java.util.Map;

/**
 * Utility class for detecting and working with Row-wrapped types.
 *
 * <p>A Row-wrapped type is a record that preserves original Excel row positions
 * for round-trip operations. It has the structure:</p>
 *
 * <pre>
 * type PersonRow record {|
 *     *xlsx:Row;        // rowIndex: int
 *     Person? value;    // nullable inner type
 * |};
 * </pre>
 */
public final class RowTypeUtils {

    /** Field name for the row index in Row-wrapped types. */
    public static final String ROW_INDEX_FIELD = "rowIndex";

    /** Field name for the value in Row-wrapped types. */
    public static final String VALUE_FIELD = "value";

    private RowTypeUtils() {
        // Private constructor to prevent instantiation
    }

    /**
     * Check if a RecordType is a Row-wrapped type.
     *
     * <p>A Row-wrapped type must have exactly 2 fields:</p>
     * <ul>
     *   <li>{@code rowIndex} - must be of type {@code int}</li>
     *   <li>{@code value} - the inner data type (usually nullable record)</li>
     * </ul>
     *
     * @param recordType The record type to check
     * @return true if the type is Row-wrapped, false otherwise
     */
    public static boolean isRowWrappedType(RecordType recordType) {
        if (recordType == null) {
            return false;
        }

        Map<String, Field> fields = recordType.getFields();

        // Must have exactly 2 fields
        if (fields.size() != 2) {
            return false;
        }

        // Must have rowIndex field of type int
        Field rowIndexField = fields.get(ROW_INDEX_FIELD);
        if (rowIndexField == null) {
            return false;
        }
        // Resolve referenced type before checking tag (important for module-defined types)
        Type rowIndexType = TypeUtils.getReferredType(rowIndexField.getFieldType());
        if (rowIndexType.getTag() != TypeTags.INT_TAG) {
            return false;
        }

        // Must have value field
        Field valueField = fields.get(VALUE_FIELD);
        if (valueField == null) {
            return false;
        }

        return true;
    }

    /**
     * Extract the inner value type from a Row-wrapped record type.
     *
     * <p>For a type like {@code PersonRow}, this returns the {@code Person?} type.
     * If the value is a union with null (e.g., {@code Person?}), returns the non-null member.</p>
     *
     * @param rowWrappedType The Row-wrapped record type
     * @return The inner value type (with null unwrapped if it's a nilable type),
     *         or null if not a valid Row-wrapped type
     */
    public static Type extractValueType(RecordType rowWrappedType) {
        if (!isRowWrappedType(rowWrappedType)) {
            return null;
        }

        Field valueField = rowWrappedType.getFields().get(VALUE_FIELD);
        Type valueType = valueField.getFieldType();

        // Resolve referenced type before checking tag
        Type resolvedValueType = TypeUtils.getReferredType(valueType);

        // If it's a union type (e.g., Person?), extract the non-null member
        if (resolvedValueType.getTag() == TypeTags.UNION_TAG) {
            UnionType unionType = (UnionType) resolvedValueType;
            for (Type memberType : unionType.getMemberTypes()) {
                // Resolve and check each member
                Type resolvedMemberType = TypeUtils.getReferredType(memberType);
                // Return the first non-null member
                if (resolvedMemberType.getTag() != TypeTags.NULL_TAG) {
                    return resolvedMemberType;
                }
            }
        }

        return resolvedValueType;
    }

    /**
     * Get the raw value field type from a Row-wrapped record type (without unwrapping union).
     *
     * @param rowWrappedType The Row-wrapped record type
     * @return The raw value field type (may be nullable union), or null if not valid
     */
    public static Type getRawValueType(RecordType rowWrappedType) {
        if (!isRowWrappedType(rowWrappedType)) {
            return null;
        }

        Field valueField = rowWrappedType.getFields().get(VALUE_FIELD);
        return valueField.getFieldType();
    }

    /**
     * Check if the value field of a Row-wrapped type is nullable.
     *
     * @param rowWrappedType The Row-wrapped record type
     * @return true if the value field is nullable (e.g., {@code Person?}), false otherwise
     */
    public static boolean isValueNullable(RecordType rowWrappedType) {
        if (!isRowWrappedType(rowWrappedType)) {
            return false;
        }

        Field valueField = rowWrappedType.getFields().get(VALUE_FIELD);
        Type valueType = valueField.getFieldType();

        // Resolve referenced type before checking tag
        Type resolvedValueType = TypeUtils.getReferredType(valueType);

        // Check if it's a union with null
        if (resolvedValueType.getTag() == TypeTags.UNION_TAG) {
            UnionType unionType = (UnionType) resolvedValueType;
            for (Type memberType : unionType.getMemberTypes()) {
                Type resolvedMemberType = TypeUtils.getReferredType(memberType);
                if (resolvedMemberType.getTag() == TypeTags.NULL_TAG) {
                    return true;
                }
            }
        }

        return false;
    }
}
