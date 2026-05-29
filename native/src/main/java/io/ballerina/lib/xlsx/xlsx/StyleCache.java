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

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * Per-call cell-style cache. Dedupes {@link CellStyle} creation within a single
 * write call so repeated date writes don't blow past Excel's ~64K style limit.
 *
 * <p>Lives only for the duration of one call and is GC'd with the calling write —
 * no static state, no cross-call sharing. Not thread-safe; each public write
 * entry point should instantiate a fresh {@code StyleCache} on the calling thread.</p>
 */
final class StyleCache {
    private final Workbook workbook;
    private final CreationHelper creationHelper;
    private final Map<String, CellStyle> styles = new HashMap<>();

    StyleCache(Workbook workbook) {
        this.workbook = workbook;
        this.creationHelper = workbook.getCreationHelper();
    }

    CellStyle getOrCreate(String formatString) {
        return styles.computeIfAbsent(formatString, fmt -> {
            CellStyle style = workbook.createCellStyle();
            style.setDataFormat(creationHelper.createDataFormat().getFormat(fmt));
            return style;
        });
    }
}
