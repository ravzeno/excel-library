package com.ravzeno.excel;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * Optional presentation settings shared by an {@link ExcelDataBuilder} instance.
 *
 * @param headerStyle POI style applied to cells in the header row; {@code null} keeps the generated
 *     style.
 * @param dataStyle POI style applied to cells in data rows; {@code null} keeps the generated style.
 * @param headerRowHeight header-row height in POI points; {@code null} uses the library default.
 * @param fontFamily font face used by generated header and data styles; {@code null} uses POI's
 *     default font.
 */
public record ExcelDataBuilderStyle(
    XSSFCellStyle headerStyle, XSSFCellStyle dataStyle, Float headerRowHeight, String fontFamily) {
  /** Default header height, in POI points. */
  public static final float DEFAULT_HEADER_ROW_HEIGHT = 450f;
}
