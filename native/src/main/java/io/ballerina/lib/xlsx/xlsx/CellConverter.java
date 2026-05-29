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

import io.ballerina.lib.xlsx.utils.XlsxConfig;
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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;

/**
 * Utility class for converting Excel cells to Ballerina values.
 */
public final class CellConverter {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter DATETIME_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final DateTimeFormatter TIME_FORMAT = DateTimeFormatter.ofPattern("HH:mm:ss");

    // Excel date format patterns for cell styling
    private static final String EXCEL_DATE_FORMAT = "yyyy-mm-dd";
    private static final String EXCEL_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss";
    private static final String EXCEL_TIME_FORMAT = "h:mm:ss";

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
    private static final long NANOS_PER_DAY = 86_400_000_000_000L;
    private static final long NANOS_PER_SECOND = 1_000_000_000L;

    // Excel epoch reference dates for direct serial ↔ LocalDate math.
    // We bypass POI's java.util.Date conversion (which uses LocalUtil.getLocaleCalendar
    // and thus inherits the system timezone) by computing directly in UTC.
    private static final LocalDate EPOCH_1900 = LocalDate.of(1900, 1, 1);
    private static final LocalDate EPOCH_1904 = LocalDate.of(1904, 1, 1);
    // Excel's 1900 epoch has the Lotus 1-2-3 leap-year bug: serial 60 represents
    // the non-existent "1900-02-29". Serials > 60 are shifted by one relative to
    // a normal day count. The threshold is hard-coded here for clarity.
    private static final long LOTUS_LEAP_DAY_SERIAL_1900 = 60;

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
                    // Direct serial → LocalDate / LocalTime math (UTC, honours isDate1904).
                    // We deliberately bypass POI's DataFormatter / cell.getDateCellValue()
                    // which both route through LocaleUtil and would inherit the system
                    // timezone. Branch on the serial's integer + fractional parts so
                    // datetime and time-only cells keep their time component instead of
                    // collapsing to a plain "yyyy-MM-dd".
                    double serial = cell.getNumericCellValue();
                    boolean is1904 = isWorkbookDate1904(cell);
                    boolean hasTimeFraction = serial != Math.floor(serial);
                    // 1900 epoch: serials < 1 are time-only by convention.
                    // 1904 epoch: serial 0 = 1904-01-01, so all serials >= 0 have a date part.
                    boolean hasDatePart = is1904 ? serial >= 0.0 : serial >= 1.0;

                    if (hasDatePart && hasTimeFraction) {
                        LocalDate date = convertSerialToLocalDate(serial, is1904);
                        LocalTime time = convertNumericToTime(serial);
                        return DATETIME_FORMAT.format(LocalDateTime.of(date, time));
                    }
                    if (hasDatePart) {
                        LocalDate date = convertSerialToLocalDate(serial, is1904);
                        return DATE_FORMAT.format(date);
                    }
                    LocalTime time = convertNumericToTime(serial);
                    return TIME_FORMAT.format(time);
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
                // Check TimeOfDay FIRST — POI considers time-only formats like "h:mm:ss"
                // as date formats, which would route through the date-conversion path and
                // pick up the 1900-epoch base date (causing epoch confusion for time values).
                Type effectiveTargetType = getEffectiveType(targetType);
                double numericValue = cell.getNumericCellValue();
                if (isTimeOfDayType(effectiveTargetType)) {
                    LocalTime tod = convertNumericToTime(numericValue);
                    return createTimeOfDayRecord(tod, (RecordType) effectiveTargetType);
                }

                // Date / Civil path: trigger on either a date-formatted cell (Excel says
                // "this is a date") or a date-typed target (caller says "give me a date").
                // We bypass POI's java.util.Date conversion entirely — computing
                // serial → LocalDate + LocalTime directly in UTC gives deterministic,
                // cross-machine, timezone-independent round-trips.
                if (DateUtil.isCellDateFormatted(cell)
                        || isDateType(effectiveTargetType) || isCivilType(effectiveTargetType)) {
                    boolean is1904 = isWorkbookDate1904(cell);
                    LocalDate localDate = convertSerialToLocalDate(numericValue, is1904);
                    LocalTime localTime = convertNumericToTime(numericValue);
                    return convertDate(localDate, localTime, targetType);
                }

                return convertNumeric(numericValue, targetType);

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
                    // Fall back to double parsing — accept whole-number-shaped strings
                    // like "42.0" but reject fractional values rather than silently
                    // truncating them to int.
                    try {
                        double d = Double.parseDouble(value.trim());
                        if (Double.isInfinite(d) || d != Math.floor(d)) {
                            throw new TypeConversionException(
                                    "Cannot convert '" + value + "' to int (non-integer value)",
                                    value, "int", "string");
                        }
                        return (long) d;
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
                if (Double.isInfinite(value) || value != Math.floor(value)) {
                    throw new TypeConversionException(
                            "Cannot convert " + value + " to int (non-integer value)",
                            String.valueOf(value), "int", "double");
                }
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

    /**
     * Build a Ballerina time record (time:Date / time:TimeOfDay / time:Civil) from
     * a naive {@link LocalDate} and {@link LocalTime}. Falls back to an ISO string
     * for non-time target types.
     */
    private static Object convertDate(LocalDate localDate, LocalTime localTime, Type targetType) {
        if (localDate == null) {
            return null;
        }

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
     * Read the workbook's date-system flag (1900 vs 1904 epoch). POI exposes
     * {@code isDate1904()} on {@code XSSFWorkbook} but not on the generic
     * {@code Workbook} interface; we default to false for non-XSSF formats
     * (which use the 1900 epoch).
     */
    static boolean isWorkbookDate1904(Cell cell) {
        Workbook wb = cell.getSheet().getWorkbook();
        if (wb instanceof XSSFWorkbook) {
            return ((XSSFWorkbook) wb).isDate1904();
        }
        return false;
    }

    /**
     * Convert an Excel serial number to a naive LocalDate, bypassing POI's
     * {@code DateUtil.getJavaDate} which routes through {@code java.util.Date} and
     * inherits the system timezone via {@code LocaleUtil.getLocaleCalendar}.
     *
     * <p>Excel dates are conceptually naive — the serial represents what the
     * spreadsheet shows, not an absolute instant. Computing directly from the
     * epoch gives deterministic, cross-machine round-trips and matches the
     * behavior of openpyxl / xlsxwriter / SheetJS.</p>
     *
     * <p>Handles the Excel 1900-epoch Lotus 1-2-3 leap-year quirk: serial 60 in
     * 1900-epoch files represents the non-existent "1900-02-29". We map it to
     * 1900-02-28 (the closest real date) and shift larger serials down by one
     * to align with Excel's display.</p>
     */
    static LocalDate convertSerialToLocalDate(double serial, boolean isDate1904) {
        long days = (long) Math.floor(serial);
        if (isDate1904) {
            return EPOCH_1904.plusDays(days);
        }
        if (days == LOTUS_LEAP_DAY_SERIAL_1900) {
            // Phantom 1900-02-29 — coerce to the closest real date.
            return LocalDate.of(1900, 2, 28);
        }
        if (days > LOTUS_LEAP_DAY_SERIAL_1900) {
            days -= 1;
        }
        // 1900 epoch: serial 1 = 1900-01-01, so the base is one day before.
        return EPOCH_1900.plusDays(days - 1);
    }

    /**
     * Convert a naive {@link LocalDate} (with an optional {@link LocalTime}) to
     * an Excel serial number, bypassing POI's Date-based conversion path.
     *
     * <p>Inverse of {@link #convertSerialToLocalDate}; reintroduces the Lotus
     * leap-day quirk for 1900-epoch files so the resulting serial matches
     * Excel's display.</p>
     */
    private static double convertLocalDateTimeToSerial(LocalDate date, LocalTime time, boolean isDate1904) {
        long days;
        if (isDate1904) {
            days = ChronoUnit.DAYS.between(EPOCH_1904, date);
        } else {
            // 1900 epoch: serial 1 = 1900-01-01.
            days = ChronoUnit.DAYS.between(EPOCH_1900, date) + 1;
            // Reintroduce the phantom 1900-02-29 so serials > 59 land where Excel expects.
            if (days >= LOTUS_LEAP_DAY_SERIAL_1900) {
                days += 1;
            }
        }
        if (time == null) {
            return (double) days;
        }
        // Day fraction via BigDecimal to preserve sub-second precision before the final cast.
        BigDecimal dayFraction = BigDecimal.valueOf(time.toNanoOfDay())
                .divide(BigDecimal.valueOf(NANOS_PER_DAY), 18, RoundingMode.HALF_UP);
        return BigDecimal.valueOf(days).add(dayFraction).doubleValue();
    }

    /**
     * Convert an Excel numeric value (fractional day) to LocalTime.
     *
     * <p>Excel stores serials as {@code double} (~17 significant decimal digits).
     * When there's no date component (serial in {@code [0, 1)}, i.e. pure
     * time-of-day), all 17 digits land in the fraction and full nanosecond
     * precision survives the round-trip. When there's a date integer (Civil
     * values), the integer consumes digits and only ~microsecond precision
     * survives — the fraction picks up sub-microsecond noise that would
     * otherwise cascade into spurious minute/second drift. We round at the
     * appropriate scale so an exact "14:30:00" wall-clock value round-trips
     * cleanly rather than emerging as "14:29:59.999_999_981".</p>
     */
    private static LocalTime convertNumericToTime(double numericValue) {
        double timeFraction = numericValue - Math.floor(numericValue);
        // Scale 0 → round to nano; scale -3 → round to nearest 1000ns (microsecond).
        int roundingScale = Math.floor(Math.abs(numericValue)) >= 1.0 ? -3 : 0;
        BigDecimal nanos = BigDecimal.valueOf(timeFraction)
                .multiply(BigDecimal.valueOf(NANOS_PER_DAY))
                .setScale(roundingScale, RoundingMode.HALF_UP);
        long nanoOfDay = nanos.longValueExact();
        // Defensive clamp: rounding may land on NANOS_PER_DAY (one tick past end-of-day)
        // for values extremely close to 1.0. LocalTime.ofNanoOfDay requires [0, NANOS_PER_DAY).
        if (nanoOfDay < 0) {
            nanoOfDay = 0;
        } else if (nanoOfDay >= NANOS_PER_DAY) {
            nanoOfDay = NANOS_PER_DAY - 1;
        }
        return LocalTime.ofNanoOfDay(nanoOfDay);
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
        String lower = value.trim().toLowerCase();
        if ("true".equals(lower) || "yes".equals(lower) || "1".equals(lower)) {
            return true;
        }
        if ("false".equals(lower) || "no".equals(lower) || "0".equals(lower)) {
            return false;
        }
        throw new TypeConversionException(
                "Cannot convert '" + value + "' to boolean",
                value, "boolean", "string");
    }

    /**
     * Set a cell value from a Ballerina value.
     *
     * @param cell       The cell to set
     * @param value      The Ballerina value
     * @param styleCache Per-call style cache for date/time format styles. Must not be null
     *                   when the value may resolve to a time record.
     */
    @SuppressWarnings("unchecked")
    public static void setCellValue(Cell cell, Object value, StyleCache styleCache) {
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
                // TimeOfDay: compute the day-fraction directly (no Date, no TZ).
                double excelTime = calculateExcelTime(map);
                cell.setCellValue(excelTime);
                applyDateCellStyle(cell, map, styleCache);  // Apply h:mm:ss format
            } else {
                // Date or Civil: bypass POI's cell.setCellValue(Date) path (which would
                // route through LocaleUtil.getLocaleCalendar and inherit the system TZ).
                // Compute the Excel serial directly from naive LocalDate/LocalTime in UTC.
                boolean is1904 = isWorkbookDate1904(cell);
                Double serial = extractDateSerialFromMap(map, is1904);
                if (serial != null) {
                    cell.setCellValue(serial);
                    applyDateCellStyle(cell, map, styleCache);  // Apply date format so readers recognise it
                } else {
                    // Not a recognized time record, convert to string
                    cell.setCellValue(value.toString());
                }
            }
        } else {
            // Strings starting with "=" are written verbatim as text — formula
            // authoring on write is deferred to a future xlsx:Formula wrapper.
            cell.setCellValue(value.toString());
        }
    }

    // =============================================================================
    // TIME RECORD EXTRACTION (for write path)
    // =============================================================================

    /**
     * Compute an Excel serial number from a Ballerina time record (time:Date or
     * time:Civil). Returns {@code null} if the map's field shape doesn't match a
     * recognized time record.
     *
     * <p>Computed directly in UTC without going through {@link java.util.Date} so
     * the round-trip is deterministic and timezone-independent. TimeOfDay is handled
     * separately in {@code setCellValue} via {@link #calculateExcelTime(BMap)}.</p>
     */
    private static Double extractDateSerialFromMap(BMap<BString, Object> map, boolean isDate1904) {
        boolean hasYear = map.containsKey(YEAR_FIELD);
        boolean hasMonth = map.containsKey(MONTH_FIELD);
        boolean hasDay = map.containsKey(DAY_FIELD);
        boolean hasHour = map.containsKey(HOUR_FIELD);
        boolean hasMinute = map.containsKey(MINUTE_FIELD);

        // time:Civil has both date and time fields
        if (hasYear && hasMonth && hasDay && hasHour && hasMinute) {
            return extractCivilSerial(map, isDate1904);
        }

        // time:Date has only date fields
        if (hasYear && hasMonth && hasDay && !hasHour) {
            return extractDateOnlySerial(map, isDate1904);
        }

        // Not a recognized time record (TimeOfDay is handled separately)
        return null;
    }

    private static double extractDateOnlySerial(BMap<BString, Object> map, boolean isDate1904) {
        int year = extractIntValue(map.get(YEAR_FIELD));
        int month = extractIntValue(map.get(MONTH_FIELD));
        int day = extractIntValue(map.get(DAY_FIELD));
        LocalDate localDate = LocalDate.of(year, month, day);
        return convertLocalDateTimeToSerial(localDate, null, isDate1904);
    }

    private static double extractCivilSerial(BMap<BString, Object> map, boolean isDate1904) {
        int year = extractIntValue(map.get(YEAR_FIELD));
        int month = extractIntValue(map.get(MONTH_FIELD));
        int day = extractIntValue(map.get(DAY_FIELD));
        int hour = extractIntValue(map.get(HOUR_FIELD));
        int minute = extractIntValue(map.get(MINUTE_FIELD));

        // time:Civil.second is `decimal` in the Ballerina time module; split into integer
        // second + nano-fraction so the BigDecimal precision survives the write.
        int second = 0;
        int nano = 0;
        if (map.containsKey(SECOND_FIELD)) {
            Object secVal = map.get(SECOND_FIELD);
            if (secVal instanceof io.ballerina.runtime.api.values.BDecimal) {
                BigDecimal bd = ((io.ballerina.runtime.api.values.BDecimal) secVal).decimalValue();
                second = bd.intValue();
                BigDecimal frac = bd.subtract(BigDecimal.valueOf(second));
                BigDecimal nanoBd = frac.multiply(BigDecimal.valueOf(NANOS_PER_SECOND))
                        .setScale(0, RoundingMode.HALF_UP);
                try {
                    nano = nanoBd.intValueExact();
                } catch (ArithmeticException ignored) {
                    nano = 0;
                }
                // Defensive clamp to the valid LocalTime nano range.
                if (nano < 0) {
                    nano = 0;
                } else if (nano >= NANOS_PER_SECOND) {
                    nano = (int) (NANOS_PER_SECOND - 1);
                }
            } else {
                second = extractIntValue(secVal);
            }
        }

        LocalDate localDate = LocalDate.of(year, month, day);
        LocalTime localTime = LocalTime.of(hour, minute, second, nano);
        return convertLocalDateTimeToSerial(localDate, localTime, isDate1904);
    }

    /**
     * Apply a date cell style so POI recognizes it as date-formatted on read.
     * Chooses format based on the type of time record (Date, TimeOfDay, or Civil).
     * Uses the per-call StyleCache to dedupe style creation within a single write call.
     */
    private static void applyDateCellStyle(Cell cell, BMap<BString, Object> map, StyleCache styleCache) {
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

        cell.setCellStyle(styleCache.getOrCreate(format));
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

        // time:TimeOfDay.second is `decimal`; preserve sub-second precision by computing
        // the day-fraction in BigDecimal before the final double cast.
        BigDecimal secondBd = BigDecimal.ZERO;
        if (map.containsKey(SECOND_FIELD)) {
            Object secVal = map.get(SECOND_FIELD);
            if (secVal instanceof io.ballerina.runtime.api.values.BDecimal) {
                secondBd = ((io.ballerina.runtime.api.values.BDecimal) secVal).decimalValue();
            } else {
                secondBd = BigDecimal.valueOf(extractIntValue(secVal));
            }
        }

        BigDecimal totalSeconds = BigDecimal.valueOf(hour)
                .multiply(BigDecimal.valueOf(SECONDS_PER_HOUR))
                .add(BigDecimal.valueOf(minute).multiply(BigDecimal.valueOf(SECONDS_PER_MINUTE)))
                .add(secondBd);
        BigDecimal dayFraction = totalSeconds.divide(
                BigDecimal.valueOf(SECONDS_PER_DAY), 18, RoundingMode.HALF_UP);
        return dayFraction.doubleValue();
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
