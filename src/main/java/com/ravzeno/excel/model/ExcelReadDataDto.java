package com.ravzeno.excel.model;

/**
 * Raw text read from a single spreadsheet cell.
 *
 * @param value formatted cell text; may be {@code null} for an empty or unreadable cell
 */
public record ExcelReadDataDto(String value) {}
