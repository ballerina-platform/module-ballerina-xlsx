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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/**
 * Utility class for detecting the actual used range of a sheet.
 * This prevents issues with "ghost rows" - rows that have formatting
 * but no actual data.
 */
public final class UsedRangeDetector {

    private UsedRangeDetector() {
        // Private constructor to prevent instantiation
    }

    /**
     * Detect the actual data boundaries of a sheet, excluding ghost rows.
     * This is critical for handling real-world Excel files that may have
     * formatting applied to thousands of empty rows.
     *
     * @param sheet The sheet to analyze
     * @return CellRangeAddress representing the used range, or null if empty
     */
    public static CellRangeAddress detectUsedRange(Sheet sheet) {
        if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
            return null;
        }

        int firstDataRow = -1;
        int lastDataRow = -1;
        int minCol = Integer.MAX_VALUE;
        int maxCol = -1;

        for (Row row : sheet) {
            if (row == null) {
                continue;
            }

            boolean rowHasData = false;

            for (Cell cell : row) {
                if (hasRealData(cell)) {
                    rowHasData = true;
                    minCol = Math.min(minCol, cell.getColumnIndex());
                    maxCol = Math.max(maxCol, cell.getColumnIndex());
                }
            }

            if (rowHasData) {
                if (firstDataRow == -1) {
                    firstDataRow = row.getRowNum();
                }
                lastDataRow = row.getRowNum();
            }
        }

        // No data found
        if (firstDataRow == -1) {
            return null;
        }

        return new CellRangeAddress(firstDataRow, lastDataRow, minCol, maxCol);
    }

    /**
     * Check if a cell contains real data (not just formatting).
     *
     * @param cell The cell to check
     * @return true if the cell has actual data
     */
    public static boolean hasRealData(Cell cell) {
        if (cell == null) {
            return false;
        }

        CellType cellType = cell.getCellType();

        switch (cellType) {
            case BLANK:
                return false;

            case STRING:
                String strValue = cell.getStringCellValue();
                return strValue != null && !strValue.trim().isEmpty();

            case NUMERIC:
            case BOOLEAN:
                return true;

            case FORMULA:
                // Formulas count as data
                return true;

            case ERROR:
                // Errors don't count as useful data
                return false;

            default:
                return false;
        }
    }

    /**
     * Convert a CellRangeAddress to A1 notation string.
     *
     * @param range The range to convert
     * @return A1 notation string (e.g., "A1:D50")
     */
    public static String toA1Notation(CellRangeAddress range) {
        if (range == null) {
            return "A1:A1";
        }

        String startCell = CellReference.convertNumToColString(range.getFirstColumn())
                + (range.getFirstRow() + 1);
        String endCell = CellReference.convertNumToColString(range.getLastColumn())
                + (range.getLastRow() + 1);

        return startCell + ":" + endCell;
    }

    /**
     * Get the row count of the used range.
     *
     * @param range The range
     * @return Number of rows
     */
    public static int getRowCount(CellRangeAddress range) {
        if (range == null) {
            return 0;
        }
        return range.getLastRow() - range.getFirstRow() + 1;
    }

    /**
     * Get the column count of the used range.
     *
     * @param range The range
     * @return Number of columns
     */
    public static int getColumnCount(CellRangeAddress range) {
        if (range == null) {
            return 0;
        }
        return range.getLastColumn() - range.getFirstColumn() + 1;
    }
}
