package com.ravzeno.excel.enums;

/** Supported Java date/time patterns for Excel export and import. */
public enum ExcelDateFormat {
  /** ISO calendar date, for example {@code 2026-10-21}. */
  ISO_DATE("yyyy-MM-dd"),

  /** ISO-style date and minute, for example {@code 2026-10-21 14:30}. */
  ISO_DATETIME("yyyy-MM-dd HH:mm"),

  /** Day-first date, commonly used in Thailand. */
  THAI_DATE("dd/MM/yyyy"),

  /** Day-first date and minute, commonly used in Thailand. */
  THAI_DATETIME("dd/MM/yyyy HH:mm"),

  /** Month-first four-digit-year date. */
  US_DATE("MM/dd/yyyy"),

  /** Month-first two-digit-year date. */
  US_DATE_SHORT("MM/dd/yy"),

  /** Day-first two-digit-year date. */
  TH_DATE_SHORT("dd/MM/yy"),

  /** Day, abbreviated month name, and year. */
  MONTH_NAME("dd MMM yyyy");

  /** Java {@link java.time.format.DateTimeFormatter} pattern. */
  private final String pattern;

  ExcelDateFormat(String pattern) {
    this.pattern = pattern;
  }

  /**
   * Returns the Java date/time pattern represented by this enum value.
   *
   * @return formatter pattern
   */
  public String getPattern() {
    return pattern;
  }
}
