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

package io.ballerina.lib.data.xlsx.xlsx;

import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BObject;
import io.ballerina.runtime.api.values.BString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.ref.PhantomReference;
import java.lang.ref.Reference;
import java.lang.ref.ReferenceQueue;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * Native handle for Apache POI Workbook.
 * This class wraps the POI Workbook and provides methods called from Ballerina.
 */
public final class WorkbookHandle {

    private static final String WORKBOOK_NATIVE_KEY = "workbookNative";
    private static final String SOURCE_PATH_KEY = "sourcePathNative";

    // PhantomReference-based cleanup for leaked workbooks
    private static final ReferenceQueue<BObject> REFERENCE_QUEUE = new ReferenceQueue<>();
    private static final Map<PhantomReference<BObject>, Workbook> LEAK_CLEANUP_MAP =
            Collections.synchronizedMap(new HashMap<>());
    // Thread-safety: volatile ensures visibility across threads.
    // Combined with synchronized block in ensureCleanupThread(), this implements
    // the correct double-checked locking pattern (safe in Java 5+).
    private static volatile boolean cleanupThreadStarted = false;

    private WorkbookHandle() {
        // Private constructor to prevent instantiation
    }

    /**
     * Ensures the cleanup thread is running.
     * The thread closes any POI Workbooks whose Ballerina BObject became unreachable
     * without having close() called. This is a defensive safeguard against resource leaks.
     */
    private static void ensureCleanupThread() {
        if (!cleanupThreadStarted) {
            synchronized (WorkbookHandle.class) {
                if (!cleanupThreadStarted) {
                    Thread cleanupThread = new Thread(() -> {
                        while (true) {
                            try {
                                Reference<?> ref = REFERENCE_QUEUE.remove(); // Blocks until reference available
                                Workbook wb = LEAK_CLEANUP_MAP.remove(ref);
                                if (wb != null) {
                                    try {
                                        wb.close();
                                    } catch (IOException ignored) {
                                        // Best effort cleanup
                                    }
                                }
                            } catch (InterruptedException e) {
                                Thread.currentThread().interrupt();
                                break;
                            }
                        }
                    }, "xlsx-workbook-cleanup");
                    cleanupThread.setDaemon(true);
                    cleanupThread.start();
                    cleanupThreadStarted = true;
                }
            }
        }
    }

    /**
     * Registers a workbook for cleanup if the Ballerina object becomes unreachable.
     */
    private static void registerForCleanup(BObject workbookObj, Workbook workbook) {
        ensureCleanupThread();
        PhantomReference<BObject> ref = new PhantomReference<>(workbookObj, REFERENCE_QUEUE);
        LEAK_CLEANUP_MAP.put(ref, workbook);
    }

    /**
     * Unregisters a workbook from cleanup tracking (called when properly closed).
     */
    private static void unregisterFromCleanup(Workbook workbook) {
        // Remove from leak cleanup tracking since it was closed properly
        LEAK_CLEANUP_MAP.values().remove(workbook);
    }

    /**
     * Open a workbook directly from a file path.
     *
     * @param workbookObj Ballerina Workbook object
     * @param filePath    Path to the XLSX file
     * @return null on success, error on failure
     */
    public static Object openWorkbookFromPath(BObject workbookObj, BString filePath) {
        try (FileInputStream fis = new FileInputStream(filePath.getValue())) {
            Workbook workbook = WorkbookFactory.create(fis);
            if (workbook == null) {
                return DiagnosticLog.error("Failed to create workbook from file: " + filePath.getValue());
            }
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            workbookObj.addNativeData(SOURCE_PATH_KEY, filePath.getValue());
            registerForCleanup(workbookObj, workbook);
            return null;
        } catch (IOException e) {
            return DiagnosticLog.fileNotFoundError("Failed to open workbook: " + filePath.getValue(), e);
        } catch (Exception e) {
            return DiagnosticLog.error("Error opening workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Create a new empty workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return null on success, error on failure
     */
    public static Object createNewWorkbook(BObject workbookObj) {
        try {
            Workbook workbook = new XSSFWorkbook();
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            registerForCleanup(workbookObj, workbook);
            return null;
        } catch (Exception e) {
            return DiagnosticLog.error("Error creating workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Get all sheet names from workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Array of sheet names
     */
    public static BArray getSheetNames(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);
        int sheetCount = workbook.getNumberOfSheets();

        BString[] names = new BString[sheetCount];
        for (int i = 0; i < sheetCount; i++) {
            names[i] = StringUtils.fromString(workbook.getSheetName(i));
        }

        return ValueCreator.createArrayValue(names);
    }

    /**
     * Get number of sheets in workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Sheet count
     */
    public static long getSheetCount(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);
        return workbook.getNumberOfSheets();
    }

    /**
     * Get a sheet by name.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param name        Sheet name
     * @return null on success, error if not found
     */
    public static Object getSheet(BObject workbookObj, BObject sheetObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);
        String sheetName = name.getValue();

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return DiagnosticLog.sheetNotFoundError(sheetName);
        }

        SheetHandle.initSheet(sheetObj, sheet);
        return null;
    }

    /**
     * Get a sheet by index.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param index       Sheet index (0-based)
     * @return null on success, error if index out of range
     */
    public static Object getSheetByIndex(BObject workbookObj, BObject sheetObj, long index) {
        Workbook workbook = getWorkbook(workbookObj);
        int idx = (int) index;

        if (idx < 0 || idx >= workbook.getNumberOfSheets()) {
            return DiagnosticLog.sheetNotFoundError(idx, workbook.getNumberOfSheets() - 1);
        }

        Sheet sheet = workbook.getSheetAt(idx);
        SheetHandle.initSheet(sheetObj, sheet);
        return null;
    }

    /**
     * Create a new sheet in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param name        Name for the new sheet
     * @return null on success, error on failure
     */
    public static Object createSheet(BObject workbookObj, BObject sheetObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);
        String sheetName = name.getValue();

        try {
            // Check if sheet already exists
            if (workbook.getSheet(sheetName) != null) {
                return DiagnosticLog.error("Sheet '" + sheetName + "' already exists");
            }

            Sheet sheet = workbook.createSheet(sheetName);
            SheetHandle.initSheet(sheetObj, sheet);
            return null;

        } catch (Exception e) {
            return DiagnosticLog.error("Error creating sheet: " + e.getMessage(), e);
        }
    }

    /**
     * Save workbook to its source file (overwrites original).
     * Returns error if workbook has no associated source path.
     *
     * @param workbookObj Ballerina Workbook object
     * @return null on success, error on failure
     */
    public static Object saveToSource(BObject workbookObj) {
        String sourcePath = getSourcePath(workbookObj);
        if (sourcePath == null) {
            return DiagnosticLog.error(
                    "Cannot save: workbook has no source file. Use saveAs(path) instead.");
        }
        return saveToPath(workbookObj, StringUtils.fromString(sourcePath));
    }

    /**
     * Save workbook to a specified path.
     *
     * @param workbookObj Ballerina Workbook object
     * @param filePath    Path to save the file
     * @return null on success, error on failure
     */
    public static Object saveToPath(BObject workbookObj, BString filePath) {
        Workbook workbook = getWorkbook(workbookObj);

        try (FileOutputStream fos = new FileOutputStream(filePath.getValue())) {
            workbook.write(fos);
            return null;
        } catch (IOException e) {
            return DiagnosticLog.error("Failed to save workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Set the source path for the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param filePath    Path to associate with the workbook
     * @return null
     */
    public static Object setSourcePath(BObject workbookObj, BString filePath) {
        workbookObj.addNativeData(SOURCE_PATH_KEY, filePath.getValue());
        return null;
    }

    /**
     * Delete a sheet by name.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Name of the sheet to delete
     * @return null on success, error if sheet not found
     */
    public static Object deleteSheet(BObject workbookObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);
        String sheetName = name.getValue();
        int index = workbook.getSheetIndex(sheetName);

        if (index == -1) {
            return DiagnosticLog.sheetNotFoundError(sheetName);
        }

        workbook.removeSheetAt(index);
        return null;
    }

    /**
     * Delete a sheet by index.
     *
     * @param workbookObj Ballerina Workbook object
     * @param index       Index of the sheet to delete (0-based)
     * @return null on success, error if index out of range
     */
    public static Object deleteSheetByIndex(BObject workbookObj, long index) {
        Workbook workbook = getWorkbook(workbookObj);
        int idx = (int) index;

        if (idx < 0 || idx >= workbook.getNumberOfSheets()) {
            return DiagnosticLog.sheetNotFoundError(idx, workbook.getNumberOfSheets() - 1);
        }

        workbook.removeSheetAt(idx);
        return null;
    }

    /**
     * Close the workbook and release resources.
     *
     * @param workbookObj Ballerina Workbook object
     * @return null on success, error on failure
     */
    public static Object close(BObject workbookObj) {
        Workbook workbook = (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
        if (workbook != null) {
            try {
                unregisterFromCleanup(workbook);
                workbook.close();
                workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, null);
            } catch (IOException e) {
                return DiagnosticLog.error("Error closing workbook: " + e.getMessage(), e);
            }
        }
        return null;
    }

    // =============================================================================
    // TABLE METHODS
    // =============================================================================

    /**
     * Get a table by name from anywhere in the workbook.
     * Table names are unique across the entire workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Table name
     * @return Ballerina Table object or error
     */
    public static Object getTable(BObject workbookObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);

        if (!(workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook)) {
            return DiagnosticLog.error("Tables are only supported in XLSX format");
        }

        org.apache.poi.xssf.usermodel.XSSFWorkbook xssfWorkbook =
                (org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook;
        String tableName = name.getValue();

        // Search all sheets for the table
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            org.apache.poi.xssf.usermodel.XSSFSheet sheet =
                    (org.apache.poi.xssf.usermodel.XSSFSheet) xssfWorkbook.getSheetAt(i);
            for (org.apache.poi.xssf.usermodel.XSSFTable table : sheet.getTables()) {
                if (tableName.equals(table.getName()) || tableName.equals(table.getDisplayName())) {
                    return createBallerinaTable(table, sheet);
                }
            }
        }

        return DiagnosticLog.tableNotFoundError(tableName);
    }

    /**
     * Get all tables from all sheets in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Array of Ballerina Table objects
     */
    public static BArray getAllTables(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);

        // Create array with proper Table type instead of anydata
        io.ballerina.runtime.api.types.Type tableType = ValueCreator.createObjectValue(
                io.ballerina.lib.data.xlsx.utils.ModuleUtils.getModule(),
                io.ballerina.lib.data.xlsx.utils.Constants.TABLE_TYPE).getType();
        io.ballerina.runtime.api.types.ArrayType tableArrayType =
                io.ballerina.runtime.api.creators.TypeCreator.createArrayType(tableType);

        if (!(workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook)) {
            // Return empty array for non-XLSX workbooks
            return ValueCreator.createArrayValue(tableArrayType);
        }

        org.apache.poi.xssf.usermodel.XSSFWorkbook xssfWorkbook =
                (org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook;

        BArray result = ValueCreator.createArrayValue(tableArrayType);

        // Iterate all sheets and collect all tables
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            org.apache.poi.xssf.usermodel.XSSFSheet sheet =
                    (org.apache.poi.xssf.usermodel.XSSFSheet) xssfWorkbook.getSheetAt(i);
            for (org.apache.poi.xssf.usermodel.XSSFTable table : sheet.getTables()) {
                result.append(createBallerinaTable(table, sheet));
            }
        }

        return result;
    }

    /**
     * Create a Ballerina Table object from POI XSSFTable.
     */
    private static BObject createBallerinaTable(org.apache.poi.xssf.usermodel.XSSFTable table,
                                                 org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
        BObject tableObj = ValueCreator.createObjectValue(
                io.ballerina.lib.data.xlsx.utils.ModuleUtils.getModule(),
                io.ballerina.lib.data.xlsx.utils.Constants.TABLE_TYPE);
        TableHandle.initTable(tableObj, table, sheet);
        return tableObj;
    }

    /**
     * Get the native Workbook from Ballerina object.
     */
    static Workbook getWorkbook(BObject workbookObj) {
        return (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
    }

    /**
     * Get the source path associated with the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Source path, or null if not set
     */
    private static String getSourcePath(BObject workbookObj) {
        return (String) workbookObj.getNativeData(SOURCE_PATH_KEY);
    }
}
