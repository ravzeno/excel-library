package com.ravzeno.excel.model.import_;

/**
 * Location settings for an import. All row and sheet numbers are one-based to match Excel's UI.
 *
 * @param sheetNo worksheet number; {@code null} defaults to the first sheet
 * @param headerRow row containing headers; {@code null} defaults to row 1
 * @param dataRow first row containing data; {@code null} defaults to row 2
 */
public record ExcelImportRequest(Integer sheetNo, Integer headerRow, Integer dataRow) {
    /**
     * Resolves the requested one-based worksheet number.
     *
     * @return requested sheet number, or 1 when unspecified
     */
    public int getSheetNoOrDefault() { return sheetNo == null ? 1 : sheetNo; }
    /**
     * Resolves the requested one-based header row number.
     *
     * @return requested header row, or 1 when unspecified
     */
    public int getHeaderRowOrDefault() { return headerRow == null ? 1 : headerRow; }
    /**
     * Resolves the requested one-based first data row number.
     *
     * @return requested first data row, or 2 when unspecified
     */
    public int getDataRowOrDefault() { return dataRow == null ? 2 : dataRow; }
}
