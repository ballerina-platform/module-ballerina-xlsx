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
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

/**
 * Builds and reads the Ballerina {@code CellRange} record (0-based row/column indices).
 *
 * Keeping the record's field names and (un)marshalling in one place avoids the same string
 * keys being repeated at every build/read site.
 */
public final class CellRangeUtils {

    private static final String RECORD_NAME = "CellRange";
    private static final String FIRST_ROW = "firstRowIndex";
    private static final String LAST_ROW = "lastRowIndex";
    private static final String FIRST_COLUMN = "firstColumnIndex";
    private static final String LAST_COLUMN = "lastColumnIndex";

    private CellRangeUtils() {
    }

    /**
     * Create a {@code CellRange} record from 0-based bounds.
     */
    public static BMap<BString, Object> create(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        BMap<BString, Object> cellRange = ValueCreator.createRecordValue(ModuleUtils.getModule(), RECORD_NAME);
        cellRange.put(StringUtils.fromString(FIRST_ROW), (long) firstRow);
        cellRange.put(StringUtils.fromString(LAST_ROW), (long) lastRow);
        cellRange.put(StringUtils.fromString(FIRST_COLUMN), (long) firstColumn);
        cellRange.put(StringUtils.fromString(LAST_COLUMN), (long) lastColumn);
        return cellRange;
    }

    public static int firstRow(BMap<BString, Object> cellRange) {
        return ((Long) cellRange.get(StringUtils.fromString(FIRST_ROW))).intValue();
    }

    public static int lastRow(BMap<BString, Object> cellRange) {
        return ((Long) cellRange.get(StringUtils.fromString(LAST_ROW))).intValue();
    }

    public static int firstColumn(BMap<BString, Object> cellRange) {
        return ((Long) cellRange.get(StringUtils.fromString(FIRST_COLUMN))).intValue();
    }

    public static int lastColumn(BMap<BString, Object> cellRange) {
        return ((Long) cellRange.get(StringUtils.fromString(LAST_COLUMN))).intValue();
    }
}
