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

import io.ballerina.lib.xlsx.utils.DiagnosticLog;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BObject;
import io.ballerina.runtime.api.values.BString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.ref.PhantomReference;
import java.lang.ref.Reference;
import java.lang.ref.ReferenceQueue;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.IdentityHashMap;
import java.util.Map;

/**
 * Native handle for Apache POI Workbook.
 * This class wraps the POI Workbook and provides methods called from Ballerina.
 */
public final class WorkbookHandle {

    // Ballerina type name for the public `Workbook` object type. Native instance construction
    // must target the concrete class name.
    private static final String WORKBOOK_TYPE = "Workbook";

    private static final String WORKBOOK_NATIVE_KEY = "workbookNative";
    private static final String SOURCE_PATH_KEY = "sourcePathNative";
    private static final String PHANTOM_REF_KEY = "phantomRef";

    // Excel's rules for sheet names: 1-31 characters, no \ / ? * [ ] :.
    private static final int MAX_SHEET_NAME_LENGTH = 31;
    private static final String INVALID_SHEET_CHARS = "\\/?*[]:";

    // Set of Sheet/Table BObjects vended from this workbook. Used to null out their
    // native data on close()/deleteSheet()/deleteTable() so that subsequent method
    // calls return a typed Error rather than touching a closed POI workbook.
    private static final String VENDED_HANDLES_KEY = "vendedHandles";

    // PhantomReference-based cleanup for leaked workbooks.
    //
    // Map direction: Workbook → PhantomReference (IdentityHashMap because Workbook
    // identity, not equals, is the comparison we want). Keyed by Workbook so
    // unregisterFromCleanup() is O(1). The matching PhantomReference is also stashed
    // on the BObject's native data so we can clear it on close without a lookup.
    private static final ReferenceQueue<BObject> REFERENCE_QUEUE = new ReferenceQueue<>();
    private static final Map<Workbook, PhantomReference<BObject>> LEAK_CLEANUP_MAP =
            new IdentityHashMap<>();
    // Reverse lookup for the cleanup thread (only path that arrives with a Reference
    // and needs to find the matching Workbook). Same identity semantics.
    private static final Map<Reference<?>, Workbook> REF_TO_WORKBOOK = new IdentityHashMap<>();
    private static final Object CLEANUP_MAP_LOCK = new Object();

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
                                Workbook wb;
                                synchronized (CLEANUP_MAP_LOCK) {
                                    wb = REF_TO_WORKBOOK.remove(ref);
                                    if (wb != null) {
                                        LEAK_CLEANUP_MAP.remove(wb);
                                    }
                                }
                                if (wb != null) {
                                    try {
                                        wb.close();
                                    } catch (IOException ignored) {
                                        // Best effort cleanup
                                    }
                                }
                            } catch (InterruptedException e) {
                                // Daemon thread; the only legitimate interrupt source would be
                                // JVM shutdown, which uses daemon-thread termination, not interrupt.
                                // Restoring the flag would cause REFERENCE_QUEUE.remove() to throw
                                // again immediately (tight loop). Breaking would permanently disable
                                // cleanup for the rest of the process. Swallow + continue keeps the
                                // safety net alive.
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
        synchronized (CLEANUP_MAP_LOCK) {
            LEAK_CLEANUP_MAP.put(workbook, ref);
            REF_TO_WORKBOOK.put(ref, workbook);
        }
        workbookObj.addNativeData(PHANTOM_REF_KEY, ref);
    }

    /**
     * Unregisters a workbook from cleanup tracking (called when properly closed).
     * O(1) — the matching PhantomReference is fetched from the BObject's native data.
     */
    private static void unregisterFromCleanup(BObject workbookObj, Workbook workbook) {
        PhantomReference<?> ref;
        synchronized (CLEANUP_MAP_LOCK) {
            ref = LEAK_CLEANUP_MAP.remove(workbook);
            if (ref != null) {
                REF_TO_WORKBOOK.remove(ref);
            }
        }
        if (ref != null) {
            ref.clear();
        }
        workbookObj.addNativeData(PHANTOM_REF_KEY, null);
    }

    /**
     * Register a Sheet or Table BObject as vended from this workbook. The handle is
     * tracked so that its native data can be nulled on close()/deleteSheet()/deleteTable().
     * Also stashes a back-reference to the parent workbook on the handle so Table
     * vending from a Sheet (Sheet.getTable/createTable/etc.) can register against the
     * correct workbook.
     */
    @SuppressWarnings("unchecked")
    static void registerVendedHandle(BObject workbookObj, BObject handleObj) {
        java.util.Set<BObject> handles =
                (java.util.Set<BObject>) workbookObj.getNativeData(VENDED_HANDLES_KEY);
        if (handles == null) {
            // Identity-based set so equals/hashCode on BObject doesn't matter.
            handles = java.util.Collections.newSetFromMap(new IdentityHashMap<>());
            workbookObj.addNativeData(VENDED_HANDLES_KEY, handles);
        }
        synchronized (handles) {
            handles.add(handleObj);
        }
        handleObj.addNativeData(SheetHandle.PARENT_WORKBOOK_KEY, workbookObj);
    }

    /**
     * Open a workbook from a file path and vend it as a Ballerina Workbook.
     *
     * @param filePath Path to the XLSX file
     * @return the Workbook object on success, error on failure
     */
    public static Object openWorkbookFromPath(BString filePath) {
        try (FileInputStream fis = new FileInputStream(filePath.getValue())) {
            Workbook workbook = WorkbookFactory.create(fis);
            if (workbook == null) {
                return DiagnosticLog.error("Failed to create workbook from file: " + filePath.getValue());
            }
            BObject workbookObj = createBallerinaWorkbook();
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            workbookObj.addNativeData(SOURCE_PATH_KEY, filePath.getValue());
            registerForCleanup(workbookObj, workbook);
            return workbookObj;
        } catch (FileNotFoundException e) {
            return DiagnosticLog.fileNotFoundError("Failed to open workbook: " + filePath.getValue(), e);
        } catch (IOException e) {
            // WorkbookFactory.create throws IOException for parse failures (corrupted ZIP,
            // malformed OOXML, encrypted files). The file existed; reading it as XLSX failed.
            return DiagnosticLog.parseError("Failed to parse workbook: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error opening workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Open a workbook from a byte array (no associated source path) and vend it as a Workbook.
     *
     * @param bytes XLSX content as a byte array
     * @return the Workbook object on success, error on failure
     */
    public static Object openWorkbookFromBytes(BArray bytes) {
        byte[] raw = bytes.getBytes();
        try (ByteArrayInputStream bis = new ByteArrayInputStream(raw)) {
            Workbook workbook = WorkbookFactory.create(bis);
            if (workbook == null) {
                return DiagnosticLog.error("Failed to create workbook from byte array");
            }
            BObject workbookObj = createBallerinaWorkbook();
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            registerForCleanup(workbookObj, workbook);
            return workbookObj;
        } catch (IOException e) {
            // WorkbookFactory.create throws IOException for parse failures (corrupted ZIP,
            // malformed OOXML, encrypted content). The bytes were readable; interpreting
            // them as XLSX failed.
            return DiagnosticLog.parseError("Failed to parse workbook bytes: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error opening workbook from bytes: " + e.getMessage(), e);
        }
    }

    /**
     * Open a workbook for in-place editing, creating a new empty one if the path is absent.
     *
     * Unlike {@link #openWorkbookFromPath}, this is BObject-free and is *not*
     * registered for phantom-reference cleanup — the caller owns the returned workbook and
     * must close it (try-with-resources). A missing file is not an error: a new empty
     * {@link XSSFWorkbook} is returned. An existing-but-unreadable file throws
     * {@link IOException}, which the caller maps to a parse error. The stream is read fully
     * by POI and closed before return, so the returned workbook holds no lock on the file —
     * safe to atomically rewrite the same path.
     *
     * @param path destination path
     * @return an open Workbook — the existing contents if the file exists, else empty
     * @throws IOException if an existing file cannot be read as XLSX (corrupt or encrypted)
     */
    static Workbook openWorkbookForEdit(String path) throws IOException {
        Path filePath = Paths.get(path);
        if (!Files.exists(filePath)) {
            return new XSSFWorkbook();
        }
        try (FileInputStream fis = new FileInputStream(filePath.toFile())) {
            return WorkbookFactory.create(fis);
        }
    }

    /**
     * Create a new empty workbook on the given Ballerina Workbook object.
     *
     * @param workbookObj Ballerina Workbook object
     */
    public static void createNewWorkbook(BObject workbookObj) {
        Workbook workbook = new XSSFWorkbook();
        workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
        registerForCleanup(workbookObj, workbook);
    }

    /**
     * Get all sheet names from workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Array of sheet names
     */
    public static Object getSheetNames(BObject workbookObj) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            int sheetCount = workbook.getNumberOfSheets();

            BString[] names = new BString[sheetCount];
            for (int i = 0; i < sheetCount; i++) {
                names[i] = StringUtils.fromString(workbook.getSheetName(i));
            }

            return ValueCreator.createArrayValue(names);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get number of sheets in workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Sheet count
     */
    public static Object getSheetCount(BObject workbookObj) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            return (long) workbook.getNumberOfSheets();
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Check whether a sheet with the given name exists in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Sheet name
     * @return true if the sheet exists, false otherwise
     */
    public static Object hasSheet(BObject workbookObj, BString name) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            return findSheetIndexCaseInsensitive(workbook, name.getValue()) != -1;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get a sheet by name.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Sheet name
     * @return Sheet BObject on success, error if not found
     */
    public static Object getSheet(BObject workbookObj, BString name) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            String sheetName = name.getValue();

            int idx = findSheetIndexCaseInsensitive(workbook, sheetName);
            if (idx == -1) {
                return DiagnosticLog.sheetNotFoundError(sheetName);
            }
            return createBallerinaSheet(workbookObj, workbook.getSheetAt(idx));
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get a sheet by index.
     *
     * @param workbookObj Ballerina Workbook object
     * @param index       Sheet index (0-based)
     * @return Sheet BObject on success, error if index out of range
     */
    public static Object getSheetByIndex(BObject workbookObj, long index) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            int sheetCount = workbook.getNumberOfSheets();
            // Bounds-check the 64-bit index before narrowing, so a huge value cannot wrap to a
            // valid int and silently address the wrong sheet.
            if (index < 0 || index >= sheetCount) {
                return DiagnosticLog.sheetNotFoundError(index, sheetCount - 1);
            }
            return createBallerinaSheet(workbookObj, workbook.getSheetAt((int) index));
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Create a new sheet in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Name for the new sheet
     * @return Sheet BObject on success, error on failure
     */
    public static Object createSheet(BObject workbookObj, BString name) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            String sheetName = name.getValue();

            validateSheetName(sheetName);

            // Check if sheet already exists (case-insensitive, matching Excel semantics)
            if (findSheetIndexCaseInsensitive(workbook, sheetName) != -1) {
                return DiagnosticLog.sheetExistsError(sheetName);
            }

            return createBallerinaSheet(workbookObj, workbook.createSheet(sheetName));

        } catch (BallerinaErrorException e) {
            return e.getBError();
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
        try {
            writeAtomically(Paths.get(filePath.getValue()), workbook);
            // Rebind the workbook's source path so subsequent save() calls
            // write to this location. Mirrors openWorkbookFromPath which
            // registers the same key after a successful open.
            workbookObj.addNativeData(SOURCE_PATH_KEY, filePath.getValue());
            return null;
        } catch (IOException e) {
            return DiagnosticLog.error("Failed to save workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Serialize the workbook to a byte array.
     *
     * @param workbookObj Ballerina Workbook object
     * @return byte array on success, error on failure
     */
    public static Object toBytes(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);
        try {
            byte[] raw = serializeWorkbook(workbook);
            return ValueCreator.createArrayValue(raw);
        } catch (IOException e) {
            return DiagnosticLog.error("Failed to serialize workbook to bytes: " + e.getMessage(), e);
        }
    }

    /**
     * Serialize a workbook to an in-memory byte array.
     * Shared helper used for byte-array output and for atomic file writes.
     */
    static byte[] serializeWorkbook(Workbook workbook) throws IOException {
        try (java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream()) {
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    /**
     * Write a workbook to a destination path atomically.
     *
     * <p>The XLSX bytes are serialized in-memory, written to a temp file in the
     * destination's parent directory, then renamed via
     * {@link Files#move(Path, Path, java.nio.file.CopyOption...)} with
     * {@link StandardCopyOption#ATOMIC_MOVE}. If any failure occurs before the
     * rename succeeds, the destination is untouched and the temp file is removed.
     * Falls back to a non-atomic move on filesystems that reject
     * {@code ATOMIC_MOVE} (extremely rare since the temp file lives in the same
     * parent directory as the destination — same filesystem in practice).</p>
     *
     * <p>This pattern guarantees that the user's existing file is never
     * truncated by a partial write — the file at {@code destination} either
     * holds the complete previous contents or the complete new contents,
     * never a half-written mix.</p>
     *
     * @param destination the final output path
     * @param workbook    the workbook to write
     * @throws IOException if serialization, temp-file write, or rename fails
     */
    public static void writeAtomically(Path destination, Workbook workbook) throws IOException {
        Path parent = destination.getParent();
        if (parent == null) {
            parent = Paths.get(".");
        }
        Path tempFile = Files.createTempFile(parent,
                destination.getFileName().toString() + ".", ".tmp");
        boolean moved = false;
        try {
            byte[] bytes = serializeWorkbook(workbook);
            Files.write(tempFile, bytes);
            try {
                Files.move(tempFile, destination,
                        StandardCopyOption.ATOMIC_MOVE,
                        StandardCopyOption.REPLACE_EXISTING);
            } catch (AtomicMoveNotSupportedException e) {
                // Defensive fallback. Same-parent temp file means same-filesystem in
                // practice, but exotic mounts (overlayfs, NFS) can still trip this.
                Files.move(tempFile, destination, StandardCopyOption.REPLACE_EXISTING);
            }
            moved = true;
        } finally {
            if (!moved) {
                try {
                    Files.deleteIfExists(tempFile);
                } catch (IOException ignored) {
                    // Best-effort cleanup; don't mask the original failure.
                }
            }
        }
    }

    /**
     * Delete a sheet by name.
     *
     * @param workbookObj Ballerina Workbook object
     * @param name        Name of the sheet to delete
     * @return null on success, error if sheet not found
     */
    public static Object deleteSheet(BObject workbookObj, BString name) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            String sheetName = name.getValue();
            int index = findSheetIndexCaseInsensitive(workbook, sheetName);

            if (index == -1) {
                return DiagnosticLog.sheetNotFoundError(sheetName);
            }
            if (workbook.getNumberOfSheets() == 1) {
                return DiagnosticLog.error(
                        "Cannot delete the last sheet — Excel requires at least one sheet");
            }

            Sheet doomed = workbook.getSheetAt(index);
            invalidateHandlesForSheet(workbookObj, doomed);
            workbook.removeSheetAt(index);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Delete a sheet by index.
     *
     * @param workbookObj Ballerina Workbook object
     * @param index       Index of the sheet to delete (0-based)
     * @return null on success, error if index out of range
     */
    public static Object deleteSheetByIndex(BObject workbookObj, long index) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            int sheetCount = workbook.getNumberOfSheets();
            // Bounds-check the 64-bit index before narrowing (see getSheetByIndex).
            if (index < 0 || index >= sheetCount) {
                return DiagnosticLog.sheetNotFoundError(index, sheetCount - 1);
            }
            if (sheetCount == 1) {
                return DiagnosticLog.error(
                        "Cannot delete the last sheet — Excel requires at least one sheet");
            }

            int idx = (int) index;
            Sheet doomed = workbook.getSheetAt(idx);
            invalidateHandlesForSheet(workbookObj, doomed);
            workbook.removeSheetAt(idx);
            return null;
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Null native data on the vended Table handle whose POI XSSFTable matches {@code doomed}.
     * Called from {@link SheetHandle#deleteTable} so that any Table handle held by user
     * code becomes invalid after the underlying table is removed from the sheet.
     */
    @SuppressWarnings("unchecked")
    static void invalidateHandlesForTable(BObject workbookObj,
            org.apache.poi.xssf.usermodel.XSSFTable doomed) {
        java.util.Set<BObject> handles =
                (java.util.Set<BObject>) workbookObj.getNativeData(VENDED_HANDLES_KEY);
        if (handles == null) {
            return;
        }
        synchronized (handles) {
            handles.removeIf(handle -> {
                Object tableNative = handle.getNativeData(TableHandle.TABLE_NATIVE_KEY);
                if (tableNative == doomed) {
                    handle.addNativeData(TableHandle.TABLE_NATIVE_KEY, null);
                    handle.addNativeData(SheetHandle.SHEET_NATIVE_KEY, null);
                    handle.addNativeData(SheetHandle.PARENT_WORKBOOK_KEY, null);
                    return true;
                }
                return false;
            });
        }
    }

    /**
     * Null native data on every vended handle whose POI sheet matches {@code doomed}.
     * For Sheet handles, the POI Sheet is compared by identity. For Table handles, the
     * sheet containing the table is compared (cascade — deleting a sheet invalidates
     * any tables vended from it).
     */
    @SuppressWarnings("unchecked")
    private static void invalidateHandlesForSheet(BObject workbookObj, Sheet doomed) {
        java.util.Set<BObject> handles =
                (java.util.Set<BObject>) workbookObj.getNativeData(VENDED_HANDLES_KEY);
        if (handles == null) {
            return;
        }
        synchronized (handles) {
            handles.removeIf(handle -> {
                Sheet sheetNative = (Sheet) handle.getNativeData(SheetHandle.SHEET_NATIVE_KEY);
                Object tableNative = handle.getNativeData(TableHandle.TABLE_NATIVE_KEY);
                if (sheetNative == doomed) {
                    // Either a Sheet handle for `doomed`, or a Table handle on `doomed`.
                    handle.addNativeData(SheetHandle.SHEET_NATIVE_KEY, null);
                    if (tableNative != null) {
                        handle.addNativeData(TableHandle.TABLE_NATIVE_KEY, null);
                    }
                    handle.addNativeData(SheetHandle.PARENT_WORKBOOK_KEY, null);
                    return true;
                }
                return false;
            });
        }
    }

    /**
     * Flip the workbook's date-system flag (1900 vs 1904 epoch).
     *
     * <p>Package-private — reachable only from test code via a private
     * {@code @java:Method} external in {@code test_utils.bal}. The public
     * Ballerina API does not expose this; production code never reaches it
     * through any documented function.</p>
     *
     * @param workbookObj Ballerina Workbook object
     * @param flag        true for 1904 epoch, false for 1900 epoch
     * @return null on success, error on failure
     */
    public static Object setDate1904Native(BObject workbookObj, boolean flag) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            if (!(workbook instanceof XSSFWorkbook)) {
                return DiagnosticLog.error("setDate1904 requires XSSFWorkbook");
            }
            // POI 5.x doesn't expose XSSFWorkbook.setDate1904(boolean); set the flag
            // directly on the underlying CTWorkbookPr OOXML schema object.
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook ctWorkbook =
                    ((XSSFWorkbook) workbook).getCTWorkbook();
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr workbookPr =
                    ctWorkbook.getWorkbookPr();
            if (workbookPr == null) {
                workbookPr = ctWorkbook.addNewWorkbookPr();
            }
            workbookPr.setDate1904(flag);
            return null;
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting date1904 flag: " + e.getMessage(), e);
        }
    }

    /**
     * Write a formula cell at the given (row, column) on the named sheet.
     *
     * <p>Package-private — reachable only from test code via a private
     * {@code @java:Method} external in {@code test_utils.bal}. The public
     * Ballerina API does not author formulas on write (strings starting with
     * "=" are written verbatim as text); this helper exists so we can build
     * fixtures that contain real formula cells for testing
     * {@code FormulaMode.CACHED} and {@code FormulaMode.TEXT} read paths.</p>
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetName   Target sheet name
     * @param row         0-based row index
     * @param col         0-based column index
     * @param formula     Formula expression without the leading "="
     * @param cachedValue Cached numeric result stored alongside the formula, mimicking the
     *                    {@code <v>} an Excel-authored file carries (not an evaluation)
     * @return null on success, error on failure
     */
    public static Object setFormulaCellNative(BObject workbookObj, BString sheetName,
                                              long row, long col, BString formula, double cachedValue) {
        try {
            Workbook workbook = getWorkbook(workbookObj);
            Sheet sheet = workbook.getSheet(sheetName.getValue());
            if (sheet == null) {
                return DiagnosticLog.error("Sheet not found: " + sheetName.getValue());
            }
            org.apache.poi.ss.usermodel.Row poiRow = sheet.getRow((int) row);
            if (poiRow == null) {
                poiRow = sheet.createRow((int) row);
            }
            org.apache.poi.ss.usermodel.Cell cell = poiRow.getCell((int) col);
            if (cell == null) {
                cell = poiRow.createCell((int) col);
            }
            cell.setCellFormula(formula.getValue());
            // Populate the cached result the way Excel would; setting a value on a formula
            // cell updates the cached <v> while preserving the <f> formula expression.
            cell.setCellValue(cachedValue);
            return null;
        } catch (Exception e) {
            return DiagnosticLog.error("Error setting formula cell: " + e.getMessage(), e);
        }
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
            // Null native data on every vended Sheet/Table handle BEFORE closing the POI
            // workbook. Any subsequent method call on a stale handle will hit the typed-error
            // path in SheetHandle.getSheet / TableHandle.getTable rather than touching a
            // closed POI workbook.
            invalidateAllVendedHandles(workbookObj);
            try {
                unregisterFromCleanup(workbookObj, workbook);
                workbook.close();
                workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, null);
            } catch (IOException e) {
                return DiagnosticLog.error("Error closing workbook: " + e.getMessage(), e);
            }
        }
        return null;
    }

    /**
     * Null out native data on every vended Sheet/Table BObject and clear the registry.
     * Called from {@link #close(BObject)} to ensure no handle outlives the POI workbook.
     */
    @SuppressWarnings("unchecked")
    private static void invalidateAllVendedHandles(BObject workbookObj) {
        java.util.Set<BObject> handles =
                (java.util.Set<BObject>) workbookObj.getNativeData(VENDED_HANDLES_KEY);
        if (handles == null) {
            return;
        }
        synchronized (handles) {
            for (BObject handle : handles) {
                handle.addNativeData(SheetHandle.SHEET_NATIVE_KEY, null);
                handle.addNativeData(TableHandle.TABLE_NATIVE_KEY, null);
                handle.addNativeData(SheetHandle.PARENT_WORKBOOK_KEY, null);
            }
            handles.clear();
        }
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
        try {
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
                        return createBallerinaTable(workbookObj, table, sheet);
                    }
                }
            }

            return DiagnosticLog.tableNotFoundError(tableName);
        } catch (BallerinaErrorException e) {
            return e.getBError();
        }
    }

    /**
     * Get all tables from all sheets in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Array of Ballerina Table objects on success, BError on failure
     */
    public static Object getAllTables(BObject workbookObj) {
        try {
            Workbook workbook = getWorkbook(workbookObj);

            // Create array with proper Table type instead of anydata
            io.ballerina.runtime.api.types.Type tableType =
                    io.ballerina.runtime.api.utils.TypeUtils.getType(
                            ValueCreator.createObjectValue(
                                    io.ballerina.lib.xlsx.utils.ModuleUtils.getModule(),
                                    TableHandle.TABLE_TYPE));
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
                    result.append(createBallerinaTable(workbookObj, table, sheet));
                }
            }

            return result;
        } catch (Exception e) {
            return DiagnosticLog.error("Error retrieving tables: " + e.getMessage(), e);
        }
    }

    /**
     * Create a Ballerina Table object from POI XSSFTable and register it as vended
     * from {@code workbookObj} so it can be invalidated on close()/deleteSheet()/deleteTable().
     */
    private static BObject createBallerinaTable(BObject workbookObj,
                                                 org.apache.poi.xssf.usermodel.XSSFTable table,
                                                 org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
        BObject tableObj = ValueCreator.createObjectValue(
                io.ballerina.lib.xlsx.utils.ModuleUtils.getModule(),
                TableHandle.TABLE_TYPE);
        TableHandle.initTable(tableObj, table, sheet);
        registerVendedHandle(workbookObj, tableObj);
        return tableObj;
    }

    /**
     * Create a Ballerina Sheet object from a POI Sheet and register it as vended
     * from {@code workbookObj} so it can be invalidated on close()/deleteSheet().
     * Mirrors {@link #createBallerinaTable}.
     */
    private static BObject createBallerinaSheet(BObject workbookObj, Sheet sheet) {
        BObject sheetObj = ValueCreator.createObjectValue(
                io.ballerina.lib.xlsx.utils.ModuleUtils.getModule(),
                SheetHandle.SHEET_TYPE);
        SheetHandle.initSheet(sheetObj, sheet);
        registerVendedHandle(workbookObj, sheetObj);
        return sheetObj;
    }

    /**
     * Vend an empty Workbook BObject for the loading factories. Its `init` allocates a
     * transient empty POI workbook; we dispose that here so a freshly-loaded workbook can
     * take the native-data slot without orphaning the throwaway one. Mirrors
     * {@link #createBallerinaSheet}.
     */
    private static BObject createBallerinaWorkbook() {
        BObject workbookObj = ValueCreator.createObjectValue(
                io.ballerina.lib.xlsx.utils.ModuleUtils.getModule(),
                WORKBOOK_TYPE);
        Workbook transientEmpty = (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
        if (transientEmpty != null) {
            unregisterFromCleanup(workbookObj, transientEmpty);
            try {
                transientEmpty.close();
            } catch (IOException ignored) {
                // An empty in-memory workbook holds no real resources; close failure is moot.
            }
        }
        return workbookObj;
    }

    /**
     * Get the native Workbook from Ballerina object.
     */
    static Workbook getWorkbook(BObject workbookObj) {
        Workbook workbook = (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
        if (workbook == null) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Workbook handle is no longer valid. The workbook may have been closed."));
        }
        return workbook;
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

    /**
     * Validate a sheet name against Excel's rules: 1-31 characters, no
     * {@code \ / ? * [ ] :} characters. Throws {@link BallerinaErrorException}
     * with a typed Error if invalid.
     */
    static void validateSheetName(String name) {
        if (name == null || name.isEmpty()) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Sheet name cannot be empty"));
        }
        if (name.length() > MAX_SHEET_NAME_LENGTH) {
            throw new BallerinaErrorException(DiagnosticLog.error(
                    "Sheet name '" + name + "' exceeds Excel's "
                            + MAX_SHEET_NAME_LENGTH + "-character limit"));
        }
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);
            if (INVALID_SHEET_CHARS.indexOf(c) >= 0) {
                throw new BallerinaErrorException(DiagnosticLog.error(
                        "Sheet name '" + name + "' contains invalid character '"
                                + c + "'. Forbidden: \\ / ? * [ ] :"));
            }
        }
    }

    /**
     * Find a sheet index by case-insensitive name match. Excel sheet names are
     * case-insensitive on lookup (you can have either "Sales" or "sales" but
     * not both, and either lookup matches the existing sheet). Returns -1 if
     * no match.
     */
    static int findSheetIndexCaseInsensitive(Workbook workbook, String name) {
        int n = workbook.getNumberOfSheets();
        for (int i = 0; i < n; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase(name)) {
                return i;
            }
        }
        return -1;
    }
}
