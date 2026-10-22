package com.ravzeno.excel.model.export;

import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Optional header-only definition for manually created worksheets.
 *
 * @param name header text
 * @param width width in POI units (1/256th of a character)
 * @param headerColor optional POI fill colour
 */
public record ExcelHeaderColumn(String name, int width, XSSFColor headerColor) {
    /** Number of POI width units in one Excel character-width unit. */
    public static final int WPC = 256;
    /**
     * Creates a header column with caller-selected width and fill colour.
     *
     * @param name header text
     * @param width width in POI units
     * @param color optional POI fill colour
     * @return header definition
     */
    public static ExcelHeaderColumn create(String name, int width, XSSFColor color) {
        return new ExcelHeaderColumn(name, width, color);
    }
}
