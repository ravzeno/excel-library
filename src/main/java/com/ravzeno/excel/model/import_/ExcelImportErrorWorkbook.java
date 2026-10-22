package com.ravzeno.excel.model.import_;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Creates an upload-ready copy of an XLSX file with invalid cells visibly marked. */
public final class ExcelImportErrorWorkbook {
  private ExcelImportErrorWorkbook() {}

  /**
   * Highlights every cell-level import error and attaches its message as an Excel comment.
   * Coordinates in {@link ExcelImportError} are zero-based. Row-level errors without a column are
   * skipped because no unambiguous cell exists to mark.
   *
   * @param source original XLSX bytes
   * @param errors errors returned by the importer
   * @return a new XLSX binary with error annotations
   * @throws IOException when the source is not a readable XLSX file
   */
  public static byte[] create(byte[] source, List<ExcelImportError> errors) throws IOException {
    if (source == null) throw new NullPointerException("source must not be null");
    if (errors == null) throw new NullPointerException("errors must not be null");
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(source));
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      CreationHelper helper = workbook.getCreationHelper();
      for (ExcelImportError error : errors) {
        if (error.row() < 0 || error.columnIndex() < 0 || workbook.getNumberOfSheets() == 0)
          continue;
        var sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(error.row());
        if (row == null) row = sheet.createRow(error.row());
        Cell cell = row.getCell(error.columnIndex());
        if (cell == null) cell = row.createCell(error.columnIndex());
        CellStyle style = workbook.createCellStyle();
        if (cell.getCellStyle() != null) style.cloneStyleFrom(cell.getCellStyle());
        style.setFillForegroundColor(org.apache.poi.ss.usermodel.IndexedColors.ROSE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
        Comment comment =
            helper.createClientAnchor() == null
                ? null
                : sheet.createDrawingPatriarch().createCellComment(helper.createClientAnchor());
        if (comment != null) {
          comment.setString(helper.createRichTextString(error.message()));
          cell.setCellComment(comment);
        }
      }
      workbook.write(output);
      return output.toByteArray();
    }
  }

  /** Convenience overload for a complete import result. */
  public static byte[] create(byte[] source, ExcelImportResult<?> result) throws IOException {
    return create(source, result.errors());
  }
}
