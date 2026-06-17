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

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Row insert/delete primitives that shift existing content to preserve it, rather than
 * overwriting it. Used to resize a table's data region (and to insert into a sheet) without
 * overwriting whatever sits below.
 *
 * These move cell content only — a table's {@code ref} extent is independent of where the
 * rows physically sit, so the caller must always re-set the table area ({@code setRef} +
 * {@code updateHeaders}) after shifting.
 */
final class RowShifter {

    private RowShifter() {
    }

    /**
     * Insert {@code k} blank rows at {@code atRow}, shifting the rows at and below {@code atRow}
     * down by {@code k}. When {@code atRow} is past the last row there is nothing to shift — the
     * gap is already empty and the caller creates rows into it.
     */
    static void makeRoom(Sheet sheet, int atRow, int k) {
        if (k <= 0) {
            return;
        }
        int last = sheet.getLastRowNum();
        if (atRow <= last) {
            sheet.shiftRows(atRow, last, k);
        }
    }

    /**
     * Remove {@code k} rows starting at {@code atRow}, pulling the rows below up by {@code k} so
     * no gap is left. Mirrors {@code SheetHandle.deleteRow}'s remove-then-shift pattern.
     */
    static void removeRows(Sheet sheet, int atRow, int k) {
        if (k <= 0) {
            return;
        }
        for (int r = atRow; r < atRow + k; r++) {
            Row row = sheet.getRow(r);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
        int last = sheet.getLastRowNum();
        if (atRow + k <= last) {
            sheet.shiftRows(atRow + k, last, -k);
        }
    }
}
