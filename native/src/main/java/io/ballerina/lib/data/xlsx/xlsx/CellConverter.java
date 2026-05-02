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

import io.ballerina.lib.data.xlsx.utils.XlsxConfig;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.types.UnionType;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.WeakHashMap;

/**
 * Utility class for converting Excel cells to Ballerina values.
 */
public final class CellConverter {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    // Excel date format patterns for cell styling
    private static final String EXCEL_DATE_FORMAT = "yyyy-mm-dd";
    private static final String EXCEL_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss";
    private static final String EXCEL_TIME_FORMAT = "h:mm:ss";

    // Style cache: WeakHashMap ties cache lifetime to workbook lifecycle
    private static final Map<Workbook, Map<String, CellStyle>> STYLE_CACHE =
            Collections.synchronizedMap(new WeakHashMap<>());

    // Time module constants for type detection
    private static final String TIME_MODULE = "time";
    private static final String DATE_TYPE = "Date";
    private static final String TIME_OF_DAY_TYPE = "TimeOfDay";
    private static final String CIVIL_TYPE = "Civil";

    // Field name constants for time records
    private static final BString YEAR_FIELD = StringUtils.fromString("year");
    private static final BString MONTH_FIELD = StringUtils.fromString("month");
    private static final BString DAY_FIELD = StringUtils.fromString("day");
    private static final BString HOUR_FIELD = StringUtils.fromString("hour");
    private static final BString MINUTE_FIELD = StringUtils.fromString("minute");
    private static final BString SECOND_FIELD = StringUtils.fromString("second");

    // Time conversion constants
    private static final int SECONDS_PER_MINUTE = 60;
    private static final int MINUTES_PER_HOUR = 60;
    private static final int HOURS_PER_DAY = 24;
    private static final int SECONDS_PER_HOUR = SECONDS_PER_MINUTE * MINUTES_PER_HOUR;
    private static final int SECONDS_PER_DAY = SECONDS_PER_HOUR * HOURS_PER_DAY;
    private static final int MINUTES_PER_DAY = MINUTES_PER_HOUR * HOURS_PER_DAY;

    private CellConverter() {
        // Private constructor to prevent instantiation
    }

    /**
     * Convert an Excel cell to a Ballerina value.
     *
     * @param cell       The cell to convert
     * @param targetType The target Ballerina type (may be null for string default)
     * @param config     Parsing configuration
     * @return The converted value, or null for empty cells
     */
    public static Object convert(Cell cell, Type targetType, XlsxConfig config) {
        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();

        // Handle formula cells
        if (cellType == CellType.FORMULA) {
            return handleFormula(cell, targetType, config);
        }

        return convertByType(cell, cellType, targetType, config);
    }

    /**
     * Convert a cell value to string (for string[][] output).
     *
     * @param cell   The cell to convert
     * @param config Parsing configuration
     * @return String value
     */
    public static String convertToString(Cell cell, XlsxConfig config) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();

        // Handle formula cells
        if (cellType == CellType.FORMULA) {
            if (config != null && config.isFormulaModeText()) {
                return "=" + cell.getCellFormula();
            }
            // CACHED mode (default when config is null) - get cached result
            cellType = cell.getCachedFormulaResultType();
        }

        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return DATE_FORMAT.format(date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
                }
                double numValue = cell.getNumericCellValue();
                // Format as integer if it's a whole number
                if (numValue == Math.floor(numValue) && !Double.isInfinite(numValue)) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case BLANK:
                return "";

            case ERROR:
                return "#ERROR";

            default:
                return cell.toString();
        }
    }

    /**
     * Convert cell to string without config-dependent behavior.
     * Used for raw string output (e.g., error logging, JSON serialization).
     *
     * @param cell The cell to convert
     * @return String value
     */
    public static String convertToStringRaw(Cell cell) {
        return convertToString(cell, null);
    }

    private static Object handleFormula(Cell cell, Type targetType, XlsxConfig config) {
        // TEXT mode - return formula string
        if (config.isFormulaModeText()) {
            return StringUtils.fromString("=" + cell.getCellFormula());
        }

        // CACHED mode - return last calculated value
        CellType cachedType = cell.getCachedFormulaResultType();
        return convertByType(cell, cachedType, targetType, config);
    }

    private static Object convertByType(Cell cell, CellType cellType, Type targetType, XlsxConfig config) {
        switch (cellType) {
            case STRING:
                String strValue = cell.getStringCellValue();
                // Check for nil value (null or empty string)
                if (isNilValue(strValue)) {
                    return null;
                }
                return convertStringToTarget(strValue, targetType);

            case NUMERIC:
                // Check TimeOfDay FIRST - must bypass DateUtil.isCellDateFormatted() because
                // POI considers time-only formats like "h:mm:ss" as date formats, which causes
                // epoch/timezone issues when reading times stored with 1899 reference date
                Type effectiveTargetType = getEffectiveType(targetType);
                if (isTimeOfDayType(effectiveTargetType)) {
                    double numericValue = cell.getNumericCellValue();
                    LocalTime localTime = convertNumericToTime(numericValue);
                    return createTimeOfDayRecord(localTime, (RecordType) effectiveTargetType);
                }

                // Standard date formatting check (for Date and Civil types)
                if (DateUtil.isCellDateFormatted(cell)) {
                    return convertDate(cell.getDateCellValue(), targetType);
                }

                // Check other time types when cell isn't date-formatted
                // (handles dates written without explicit date formatting)
                if (isDateType(effectiveTargetType) || isCivilType(effectiveTargetType)) {
                    Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
                    return convertDate(date, targetType);
                }

                return convertNumeric(cell.getNumericCellValue(), targetType);

            case BOOLEAN:
                return convertBoolean(cell.getBooleanCellValue(), targetType);

            case BLANK:
                return null;

            case ERROR:
                // Return null for error cells in data context
                return null;

            default:
                return StringUtils.fromString(cell.toString());
        }
    }

    private static boolean isNilValue(String value) {
        return value == null || value.trim().isEmpty();
    }

    private static Object convertStringToTarget(String value, Type targetType) {
        if (targetType == null) {
            return StringUtils.fromString(value);
        }

        // Handle union types (e.g., int?, string?) by extracting the non-nil member type
        Type effectiveType = targetType;
        if (targetType.getTag() == TypeTags.UNION_TAG) {
            effectiveType = getNonNilType((UnionType) targetType);
            if (effectiveType == null) {
                // Union with only nil - return string
                return StringUtils.fromString(value);
            }
        }

        int typeTag = effectiveType.getTag();

        switch (typeTag) {
            case TypeTags.INT_TAG:
                try {
                    return Long.parseLong(value.trim());
                } catch (NumberFormatException e) {
                    // Try parsing as double first
                    try {
                        return (long) Double.parseDouble(value.trim());
                    } catch (NumberFormatException e2) {
                        throw new TypeConversionException(
                                "Cannot convert '" + value + "' to int",
                                value, "int", "string", e2);
                    }
                }

            case TypeTags.FLOAT_TAG:
                try {
                    return Double.parseDouble(value.trim());
                } catch (NumberFormatException e) {
                    throw new TypeConversionException(
                            "Cannot convert '" + value + "' to float",
                            value, "float", "string", e);
                }

            case TypeTags.DECIMAL_TAG:
                try {
                    return ValueCreator.createDecimalValue(new BigDecimal(value.trim()));
                } catch (NumberFormatException e) {
                    throw new TypeConversionException(
                            "Cannot convert '" + value + "' to decimal",
                            value, "decimal", "string", e);
                }

            case TypeTags.BOOLEAN_TAG:
                return parseBoolean(value);

            case TypeTags.STRING_TAG:
            default:
                return StringUtils.fromString(value);
        }
    }

    private static Object convertNumeric(double value, Type targetType) {
        if (targetType == null) {
            // Untyped: return as string for consistency with string cells
            if (value == Math.floor(value) && !Double.isInfinite(value)) {
                return StringUtils.fromString(String.valueOf((long) value));
            }
            return StringUtils.fromString(String.valueOf(value));
        }

        // Handle union types (e.g., int?, float?) by extracting the non-nil member type
        Type effectiveType = targetType;
        if (targetType.getTag() == TypeTags.UNION_TAG) {
            effectiveType = getNonNilType((UnionType) targetType);
            if (effectiveType == null) {
                // Union with only nil - return as string
                if (value == Math.floor(value) && !Double.isInfinite(value)) {
                    return StringUtils.fromString(String.valueOf((long) value));
                }
                return StringUtils.fromString(String.valueOf(value));
            }
        }

        int typeTag = effectiveType.getTag();

        switch (typeTag) {
            case TypeTags.INT_TAG:
                return (long) value;

            case TypeTags.FLOAT_TAG:
                return value;

            case TypeTags.DECIMAL_TAG:
                return ValueCreator.createDecimalValue(BigDecimal.valueOf(value));

            case TypeTags.STRING_TAG:
                if (value == Math.floor(value) && !Double.isInfinite(value)) {
                    return StringUtils.fromString(String.valueOf((long) value));
                }
                return StringUtils.fromString(String.valueOf(value));

            case TypeTags.BOOLEAN_TAG:
                return value != 0;

            default:
                return ValueCreator.createDecimalValue(BigDecimal.valueOf(value));
        }
    }

    private static Object convertBoolean(boolean value, Type targetType) {
        if (targetType == null) {
            // Untyped: return as string for consistency with string cells
            return StringUtils.fromString(String.valueOf(value));
        }

        // Handle union types (e.g., boolean?) by extracting the non-nil member type
        Type effectiveType = targetType;
        if (targetType.getTag() == TypeTags.UNION_TAG) {
            effectiveType = getNonNilType((UnionType) targetType);
            if (effectiveType == null) {
                // Union with only nil - return as string
                return StringUtils.fromString(String.valueOf(value));
            }
        }

        int typeTag = effectiveType.getTag();

        switch (typeTag) {
            case TypeTags.BOOLEAN_TAG:
                return value;

            case TypeTags.STRING_TAG:
                return StringUtils.fromString(String.valueOf(value));

            case TypeTags.INT_TAG:
                return value ? 1L : 0L;

            case TypeTags.FLOAT_TAG:
                return value ? 1.0 : 0.0;

            default:
                return value;
        }
    }

    private static Object convertDate(Date date, Type targetType) {
        // Handle null date (can happen for edge cases in DateUtil.getJavaDate)
        if (date == null) {
            return null;
        }

        // Convert java.util.Date to java.time types
        ZonedDateTime zdt = date.toInstant().atZone(ZoneId.systemDefault());
        LocalDate localDate = zdt.toLocalDate();
        LocalTime localTime = zdt.toLocalTime();

        // Get effective type (handle nullable types like time:Date?)
        Type effectiveType = getEffectiveType(targetType);

        // Check for Ballerina time module types
        if (isDateType(effectiveType)) {
            return createDateRecord(localDate, (RecordType) effectiveType);
        } else if (isTimeOfDayType(effectiveType)) {
            return createTimeOfDayRecord(localTime, (RecordType) effectiveType);
        } else if (isCivilType(effectiveType)) {
            return createCivilRecord(localDate, localTime, (RecordType) effectiveType);
        }

        // Fallback: ISO format string (backward compatible)
        return StringUtils.fromString(DATE_FORMAT.format(localDate));
    }

    /**
     * Get the effective type, resolving type references and handling union types.
     * For example: time:Date (TypeReferenceType) -> RecordType
     *              time:Date? (UnionType containing TypeReferenceType) -> RecordType
     */
    private static Type getEffectiveType(Type type) {
        if (type == null) {
            return null;
        }
        // Resolve type references first (e.g., time:Date is wrapped in TypeReferenceType)
        Type resolved = TypeUtils.getReferredType(type);

        // Then handle union types (e.g., time:Date?)
        if (resolved.getTag() == TypeTags.UNION_TAG) {
            Type nonNilType = getNonNilType((UnionType) resolved);
            // The extracted non-nil type may also be a TypeReferenceType, so resolve it too
            return nonNilType != null ? TypeUtils.getReferredType(nonNilType) : null;
        }
        return resolved;
    }

    /**
     * Convert an Excel numeric value (fractional day) to LocalTime.
     * In Excel, time is stored as a fraction of a day (e.g., 0.5 = 12:00 noon).
     */
    private static LocalTime convertNumericToTime(double numericValue) {
        // Excel stores time as fraction of day (0.0 to ~0.9999...)
        // Fractional part represents time, integer part represents days
        double timeFraction = numericValue - Math.floor(numericValue);
        long totalSeconds = Math.round(timeFraction * SECONDS_PER_DAY);
        int hours = (int) ((totalSeconds / SECONDS_PER_HOUR) % HOURS_PER_DAY);
        int minutes = (int) ((totalSeconds % SECONDS_PER_HOUR) / SECONDS_PER_MINUTE);
        int seconds = (int) (totalSeconds % SECONDS_PER_MINUTE);
        return LocalTime.of(hours, minutes, seconds);
    }

    // =============================================================================
    // TIME TYPE DETECTION
    // =============================================================================

    /**
     * Check if a type is from the Ballerina time module with a specific type name.
     */
    private static boolean isTimeModuleType(Type type, String typeName) {
        if (type == null || type.getTag() != TypeTags.RECORD_TYPE_TAG) {
            return false;
        }
        RecordType recordType = (RecordType) type;
        // Check module name from the package
        io.ballerina.runtime.api.Module pkg = recordType.getPackage();
        return pkg != null
                && TIME_MODULE.equals(pkg.getName())
                && typeName.equals(type.getName());
    }

    private static boolean isDateType(Type type) {
        return isTimeModuleType(type, DATE_TYPE);
    }

    private static boolean isTimeOfDayType(Type type) {
        return isTimeModuleType(type, TIME_OF_DAY_TYPE);
    }

    private static boolean isCivilType(Type type) {
        return isTimeModuleType(type, CIVIL_TYPE);
    }

    // =============================================================================
    // TIME RECORD CREATION
    // =============================================================================

    /**
     * Create a Ballerina time:Date record from LocalDate.
     * time:Date = { year: int, month: int, day: int }
     */
    private static BMap<BString, Object> createDateRecord(LocalDate date, RecordType dateType) {
        BMap<BString, Object> dateRecord = ValueCreator.createRecordValue(dateType);
        dateRecord.put(YEAR_FIELD, (long) date.getYear());
        dateRecord.put(MONTH_FIELD, (long) date.getMonthValue());
        dateRecord.put(DAY_FIELD, (long) date.getDayOfMonth());
        return dateRecord;
    }

    /**
     * Create a Ballerina time:TimeOfDay record from LocalTime.
     * time:TimeOfDay = { hour: int, minute: int, second: decimal }
     */
    private static BMap<BString, Object> createTimeOfDayRecord(LocalTime time, RecordType timeOfDayType) {
        BMap<BString, Object> timeRecord = ValueCreator.createRecordValue(timeOfDayType);
        timeRecord.put(HOUR_FIELD, (long) time.getHour());
        timeRecord.put(MINUTE_FIELD, (long) time.getMinute());
        // Ballerina uses decimal for seconds (includes fractional part)
        BigDecimal seconds = BigDecimal.valueOf(time.getSecond())
                .add(BigDecimal.valueOf(time.getNano(), 9));
        timeRecord.put(SECOND_FIELD, ValueCreator.createDecimalValue(seconds));
        return timeRecord;
    }

    /**
     * Create a Ballerina time:Civil record from LocalDate and LocalTime.
     * time:Civil = { year, month, day, hour, minute, second }
     */
    private static BMap<BString, Object> createCivilRecord(LocalDate date, LocalTime time, RecordType civilType) {
        BMap<BString, Object> civilRecord = ValueCreator.createRecordValue(civilType);
        // Date fields
        civilRecord.put(YEAR_FIELD, (long) date.getYear());
        civilRecord.put(MONTH_FIELD, (long) date.getMonthValue());
        civilRecord.put(DAY_FIELD, (long) date.getDayOfMonth());
        // Time fields
        civilRecord.put(HOUR_FIELD, (long) time.getHour());
        civilRecord.put(MINUTE_FIELD, (long) time.getMinute());
        BigDecimal seconds = BigDecimal.valueOf(time.getSecond())
                .add(BigDecimal.valueOf(time.getNano(), 9));
        civilRecord.put(SECOND_FIELD, ValueCreator.createDecimalValue(seconds));
        return civilRecord;
    }

    private static boolean parseBoolean(String value) {
        if (value == null) {
            return false;
        }
        String lower = value.trim().toLowerCase();
        return "true".equals(lower) || "yes".equals(lower) || "1".equals(lower);
    }

    /**
     * Set a cell value from a Ballerina value.
     *
     * @param cell  The cell to set
     * @param value The Ballerina value
     */
    @SuppressWarnings("unchecked")
    public static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof io.ballerina.runtime.api.values.BDecimal) {
            cell.setCellValue(((io.ballerina.runtime.api.values.BDecimal) value).decimalValue().doubleValue());
        } else if (value instanceof BMap) {
            // Check if it's a time record (Date, TimeOfDay, or Civil)
            BMap<BString, Object> map = (BMap<BString, Object>) value;

            // Check if it's a time-only record (TimeOfDay)
            boolean hasHour = map.containsKey(HOUR_FIELD);
            boolean hasYear = map.containsKey(YEAR_FIELD);

            if (hasHour && !hasYear) {
                // TimeOfDay: use direct calculation to avoid POI's pre-1900 date rejection
                // and timezone issues. Excel stores time as fraction of day.
                double excelTime = calculateExcelTime(map);
                cell.setCellValue(excelTime);
                applyDateCellStyle(cell, map);  // Apply h:mm:ss format
            } else {
                // Date or Civil: use existing Date-based approach
                Date dateValue = extractDateFromMap(map);
                if (dateValue != null) {
                    cell.setCellValue(dateValue);
                    applyDateCellStyle(cell, map);  // Apply date format so POI recognizes it on read
                } else {
                    // Not a recognized time record, convert to string
                    cell.setCellValue(value.toString());
                }
            }
        } else {
            // Default: convert to string
            String strValue = value.toString();
            // Check if this is a formula (starts with "=")
            if (strValue.startsWith("=") && strValue.length() > 1) {
                // Set as formula, stripping the leading "="
                cell.setCellFormula(strValue.substring(1));
            } else {
                cell.setCellValue(strValue);
            }
        }
    }

    // =============================================================================
    // TIME RECORD EXTRACTION (for write path)
    // =============================================================================

    /**
     * Extract a java.util.Date from a Ballerina time record (Date or Civil).
     * Returns null if the map is not a recognized time record.
     *
     * <p>Note: TimeOfDay is handled separately in setCellValue() using calculateExcelTime()
     * to avoid POI's pre-1900 date rejection and timezone issues.</p>
     */
    private static Date extractDateFromMap(BMap<BString, Object> map) {
        boolean hasYear = map.containsKey(YEAR_FIELD);
        boolean hasMonth = map.containsKey(MONTH_FIELD);
        boolean hasDay = map.containsKey(DAY_FIELD);
        boolean hasHour = map.containsKey(HOUR_FIELD);
        boolean hasMinute = map.containsKey(MINUTE_FIELD);

        // time:Civil has both date and time fields
        if (hasYear && hasMonth && hasDay && hasHour && hasMinute) {
            return extractCivilDate(map);
        }

        // time:Date has only date fields
        if (hasYear && hasMonth && hasDay && !hasHour) {
            return extractDateOnlyDate(map);
        }

        // Not a recognized time record (TimeOfDay is handled separately)
        return null;
    }

    private static Date extractDateOnlyDate(BMap<BString, Object> map) {
        int year = extractIntValue(map.get(YEAR_FIELD));
        int month = extractIntValue(map.get(MONTH_FIELD));
        int day = extractIntValue(map.get(DAY_FIELD));

        LocalDate localDate = LocalDate.of(year, month, day);
        return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    private static Date extractCivilDate(BMap<BString, Object> map) {
        int year = extractIntValue(map.get(YEAR_FIELD));
        int month = extractIntValue(map.get(MONTH_FIELD));
        int day = extractIntValue(map.get(DAY_FIELD));
        int hour = extractIntValue(map.get(HOUR_FIELD));
        int minute = extractIntValue(map.get(MINUTE_FIELD));
        int second = 0;
        if (map.containsKey(SECOND_FIELD)) {
            second = extractIntValue(map.get(SECOND_FIELD));
        }

        LocalDate localDate = LocalDate.of(year, month, day);
        LocalTime localTime = LocalTime.of(hour, minute, second);
        return Date.from(localDate.atTime(localTime).atZone(ZoneId.systemDefault()).toInstant());
    }

    /**
     * Get or create a cached CellStyle for the given workbook and format.
     * Styles are cached per workbook per format to avoid hitting Excel's ~64K style limit.
     * WeakHashMap ensures cache entries are cleaned up when workbook is garbage collected.
     */
    private static synchronized CellStyle getOrCreateStyle(Workbook workbook, String format) {
        return STYLE_CACHE
                .computeIfAbsent(workbook, k -> new HashMap<>())
                .computeIfAbsent(format, f -> {
                    CellStyle style = workbook.createCellStyle();
                    style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(f));
                    return style;
                });
    }

    /**
     * Apply a date cell style so POI recognizes it as date-formatted on read.
     * Chooses format based on the type of time record (Date, TimeOfDay, or Civil).
     */
    private static void applyDateCellStyle(Cell cell, BMap<BString, Object> map) {
        Workbook workbook = cell.getSheet().getWorkbook();

        // Detect the type of time record by checking which fields are present
        boolean hasDate = map.containsKey(YEAR_FIELD);  // Date and Civil have year
        boolean hasTime = map.containsKey(HOUR_FIELD);  // TimeOfDay and Civil have hour

        String format;
        if (hasDate && hasTime) {
            format = EXCEL_DATETIME_FORMAT;  // Civil: yyyy-mm-dd hh:mm:ss
        } else if (hasDate) {
            format = EXCEL_DATE_FORMAT;      // Date: yyyy-mm-dd
        } else {
            format = EXCEL_TIME_FORMAT;      // TimeOfDay: h:mm:ss
        }

        cell.setCellStyle(getOrCreateStyle(workbook, format));
    }

    /**
     * Safely extract an int value from various Ballerina numeric types.
     */
    private static int extractIntValue(Object value) {
        if (value instanceof Long) {
            return ((Long) value).intValue();
        } else if (value instanceof Integer) {
            return (Integer) value;
        } else if (value instanceof io.ballerina.runtime.api.values.BDecimal) {
            return ((io.ballerina.runtime.api.values.BDecimal) value).decimalValue().intValue();
        } else if (value instanceof Double) {
            return ((Double) value).intValue();
        } else if (value instanceof Number) {
            return ((Number) value).intValue();
        }
        return 0;
    }

    /**
     * Calculate Excel time serial for TimeOfDay values.
     * Excel stores time as a fraction of day (0.0 to 0.9999...).
     * For example: 9:00 AM = 9/24 = 0.375, 12:30 PM = 12.5/24 ≈ 0.5208
     *
     * <p>This method bypasses Date conversion to avoid POI's rejection of pre-1900 dates
     * and timezone-related issues.</p>
     *
     * @param map The Ballerina TimeOfDay record map
     * @return Excel time serial (fraction of day)
     */
    private static double calculateExcelTime(BMap<BString, Object> map) {
        int hour = extractIntValue(map.get(HOUR_FIELD));
        int minute = extractIntValue(map.get(MINUTE_FIELD));
        int second = 0;
        if (map.containsKey(SECOND_FIELD)) {
            second = extractIntValue(map.get(SECOND_FIELD));
        }
        return hour / (double) HOURS_PER_DAY + minute / (double) MINUTES_PER_DAY + second / (double) SECONDS_PER_DAY;
    }

    /**
     * Extract the non-nil type from a union type.
     * For example, for int?, returns int. For string?, returns string.
     *
     * @param unionType The union type
     * @return The non-nil member type, or null if union only contains nil
     */
    private static Type getNonNilType(UnionType unionType) {
        for (Type memberType : unionType.getMemberTypes()) {
            if (memberType.getTag() != TypeTags.NULL_TAG) {
                return memberType;
            }
        }
        return null;
    }
}
