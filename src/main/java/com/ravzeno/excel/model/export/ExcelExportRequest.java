package com.ravzeno.excel.model.export;

import java.util.List;

/**
 * Input settings for an in-memory Excel export.
 *
 * @param data DTO rows to export; copied defensively and normalized to an empty list when {@code
 *     null}
 * @param sheetName worksheet name; blank values become {@code Sheet1}
 * @param filename suggested download/export filename; blank values become {@code export.xlsx}
 * @param headerRowHeight optional header height in POI points
 * @param fontFamily optional font for generated styles
 * @param <T> DTO type whose fields use {@code @ExcelColumn}
 */
public record ExcelExportRequest<T>(
    List<T> data,
    String sheetName,
    String filename,
    Float headerRowHeight,
    String fontFamily,
    boolean freezeHeader) {
  /** Normalizes optional user input and protects the instance from later list mutations. */
  public ExcelExportRequest {
    data = data == null ? List.of() : List.copyOf(data);
    sheetName = sheetName == null || sheetName.isBlank() ? "Sheet1" : sheetName;
    filename = filename == null || filename.isBlank() ? "export.xlsx" : filename;
  }

  /** Backward-compatible request constructor with header freezing enabled. */
  public ExcelExportRequest(
      List<T> data, String sheetName, String filename, Float headerRowHeight, String fontFamily) {
    this(data, sheetName, filename, headerRowHeight, fontFamily, true);
  }
}
