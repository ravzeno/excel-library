package com.ravzeno.excel.enums;

/** Common Excel number-format patterns for exported numeric cells. */
public enum ExcelNumberFormat {
    /** Whole number with grouping separators. */
    DECIMAL_0("#,##0"),

    /** Number with two decimal places and grouping separators. */
    DECIMAL_2("#,##0.00"),

    /** Number with four decimal places and grouping separators. */
    DECIMAL_4("#,##0.0000"),

    /** Two-decimal numeric format; currency symbol is supplied by the caller's style if needed. */
    CURRENCY("#,##0.00"),

    /** Percentage with no fractional decimal places. */
    PERCENT("#,##0%");

    /** Native Excel number-format pattern. */
    private final String pattern;

    ExcelNumberFormat(String pattern) { this.pattern = pattern; }
    /**
     * Returns the native Excel number-format pattern.
     *
     * @return number-format pattern
     */
    public String getPattern() { return pattern; }
}
