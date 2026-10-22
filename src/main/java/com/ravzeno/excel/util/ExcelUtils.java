package com.ravzeno.excel.util;

import com.ravzeno.excel.enums.ExcelDateFormat;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.regex.Pattern;

/** Common conversion and formatting helpers shared by the export and import pipelines. */
public final class ExcelUtils {
  /** Prevents construction because this class exposes only static helpers. */
  private ExcelUtils() {}

  /**
   * Parses Excel serial dates or date text and normalizes Buddhist Era years to Common Era years.
   *
   * @param value cell value represented as a serial number or text
   * @param format preferred format tried before the remaining supported formats
   * @return parsed date, or {@code null} when none of the formats match
   */
  public static LocalDate parseDate(Object value, ExcelDateFormat format) {
    if (value == null) return null;
    String text = value.toString().trim();
    if (text.matches("^\\d+(\\.0)?$"))
      return LocalDate.of(1899, 12, 30).plusDays(Long.parseLong(text.split("\\.")[0]));
    for (ExcelDateFormat candidate : orderedFormats(format)) {
      try {
        return normalizeThaiBuddhistYear(
            LocalDate.parse(text, DateTimeFormatter.ofPattern(candidate.getPattern())));
      } catch (DateTimeParseException ignored) {
      }
    }
    return null;
  }

  /** Parses date-time text, accepting date-only input as the start of that day. */
  public static LocalDateTime parseDateTime(Object value, ExcelDateFormat format) {
    if (value == null) return null;
    String text = value.toString().trim();
    for (ExcelDateFormat candidate : orderedFormats(format)) {
      try {
        return LocalDateTime.parse(text, DateTimeFormatter.ofPattern(candidate.getPattern()));
      } catch (DateTimeParseException ignored) {
      }
    }
    LocalDate date = parseDate(text, format);
    return date == null ? null : date.atStartOfDay();
  }

  /**
   * Tests whether text can be parsed as a decimal number after removing grouping commas.
   *
   * @param value candidate text
   * @return {@code true} only for a non-blank numeric value
   */
  public static boolean isNumeric(String value) {
    if (value == null || value.isBlank()) return false;
    try {
      new BigDecimal(value.replace(",", "").trim());
      return true;
    } catch (NumberFormatException e) {
      return false;
    }
  }

  /**
   * Performs a lightweight check for supported numeric and textual date shapes.
   *
   * @param value candidate text
   * @return {@code true} when the text looks like a supported date representation
   */
  public static boolean isDate(String value) {
    if (value == null || value.isBlank()) return false;
    return Pattern.matches("^\\d+(\\.0)?$", value)
        || Pattern.matches("^\\d{1,2}[/-]\\d{1,2}[/-]\\d{2,4}$", value)
        || Pattern.matches("^\\d{4}-\\d{2}-\\d{2}$", value);
  }

  /** Builds a format list with the caller preference first, followed by every fallback. */
  private static ExcelDateFormat[] orderedFormats(ExcelDateFormat preferred) {
    ExcelDateFormat[] values = ExcelDateFormat.values();
    ExcelDateFormat[] result = new ExcelDateFormat[values.length];
    result[0] = preferred;
    int index = 1;
    for (ExcelDateFormat value : values) if (value != preferred) result[index++] = value;
    return result;
  }

  /** Converts Buddhist Era years, identified by a year above 2400, to Common Era years. */
  private static LocalDate normalizeThaiBuddhistYear(LocalDate date) {
    return date.getYear() > 2400 ? date.minusYears(543) : date;
  }
}
