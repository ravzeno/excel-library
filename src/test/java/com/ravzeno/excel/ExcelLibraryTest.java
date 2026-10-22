package com.ravzeno.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.ravzeno.excel.annotation.ExcelColumn;
import com.ravzeno.excel.annotation.ExcelImportColumn;
import com.ravzeno.excel.base.AbstractExcelExportService;
import com.ravzeno.excel.base.AbstractExcelImportService;
import com.ravzeno.excel.enums.ExcelDataType;
import com.ravzeno.excel.enums.ExcelImportProblemCode;
import com.ravzeno.excel.model.export.ExcelExportResult;
import com.ravzeno.excel.model.import_.ExcelImportErrorWorkbook;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class ExcelLibraryTest {
  @Test
  void exportsAndImportsAnnotatedRows() throws IOException {
    ExcelExportResult exported =
        new ExportService()
            .export(
                List.of(
                    new ExportRow("A-01", new BigDecimal("12.50"), LocalDate.of(2026, 10, 21))));
    assertTrue(exported.bytes().length > 0);
    var result = new ImportService().importFromExcel(exported.bytes());
    assertEquals(1, result.validRows());
    assertEquals("A-01", result.data().getFirst().code);
    assertEquals(0, new BigDecimal("12.50").compareTo(result.data().getFirst().amount));
    assertEquals(LocalDate.of(2026, 10, 21), result.data().getFirst().date);
  }

  @Test
  void factoryApisAvoidDedicatedServiceSubclasses() throws IOException {
    var exporter =
        AbstractExcelExportService.of(
            ExportRow.class, "Orders", row -> new Object[] {row.code, row.amount, row.date});
    var importer = AbstractExcelImportService.of(ImportRow.class, 1, 1, 2);

    var result =
        importer.importFromExcel(
            exporter
                .export(
                    List.of(
                        new ExportRow("A-01", new BigDecimal("12.50"), LocalDate.of(2026, 10, 21))))
                .bytes());

    assertEquals(1, result.validRows());
    assertEquals("A-01", result.data().getFirst().code);
  }

  @Test
  void exportsTemplatesAndMultipleWorksheets() throws IOException {
    var exporter = new ExportService();
    var template = exporter.exportTemplate("Template", "template.xlsx");
    var workbook =
        exporter.exportWorkbook(
            List.of(
                new com.ravzeno.excel.model.export.ExcelExportRequest<>(
                    List.of(), "First", "book.xlsx", null, null),
                new com.ravzeno.excel.model.export.ExcelExportRequest<>(
                    List.of(), "Second", "book.xlsx", null, null)),
            "book.xlsx");
    try (XSSFWorkbook templateBook = new XSSFWorkbook(new ByteArrayInputStream(template.bytes()));
        XSSFWorkbook multiBook = new XSSFWorkbook(new ByteArrayInputStream(workbook.bytes()))) {
      assertEquals("Template", templateBook.getSheetAt(0).getSheetName());
      assertEquals(2, multiBook.getNumberOfSheets());
      assertEquals("Second", multiBook.getSheetAt(1).getSheetName());
    }
  }

  @Test
  void createsAnnotatedErrorWorkbook() throws IOException {
    byte[] source = workbookWithBlankCell();
    var result = new PrimitiveImportService().importFromExcel(source);
    byte[] marked = ExcelImportErrorWorkbook.create(source, result);
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(marked))) {
      assertTrue(workbook.getSheetAt(0).getRow(1).getCell(0).getCellComment() != null);
    }
  }

  @Test
  void importsDateTimesAndAppliesNumericLimits() throws IOException {
    var dateTime =
        new DateTimeImportService().importFromExcel(workbookWithText("2026-10-21 14:30"));
    var constrained = new ConstrainedImportService().importFromExcel(workbookWithText("9"));
    assertEquals(LocalDateTime.of(2026, 10, 21, 14, 30), dateTime.data().getFirst().value);
    assertEquals(ExcelImportProblemCode.MIN_VALUE, constrained.errors().getFirst().problemCode());
  }

  @Test
  void exportsStylesForNonContiguousAnnotationIndexes() throws IOException {
    ExcelExportResult exported =
        new ExportService()
            .export(
                List.of(
                    new ExportRow("A-01", new BigDecimal("12.50"), LocalDate.of(2026, 10, 21))));
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(exported.bytes()))) {
      var sheet = workbook.getSheetAt(0);
      assertEquals("Code", sheet.getRow(0).getCell(0).getStringCellValue());
      assertEquals("Amount", sheet.getRow(0).getCell(1).getStringCellValue());
      assertEquals(
          HorizontalAlignment.LEFT, sheet.getRow(1).getCell(0).getCellStyle().getAlignment());
      assertEquals("#,##0.00", sheet.getRow(1).getCell(1).getCellStyle().getDataFormatString());
    }
  }

  @Test
  void importsFormulaValues() throws IOException {
    byte[] file = workbookWithFormula();
    var result = new FormulaImportService().importFromExcel(file);
    assertEquals(1, result.validRows());
    assertEquals(0, new BigDecimal("5").compareTo(result.data().getFirst().amount));
  }

  @Test
  void rejectsBlankRequiredPrimitiveValues() throws IOException {
    var result = new PrimitiveImportService().importFromExcel(workbookWithBlankCell());
    assertEquals(0, result.validRows());
    assertEquals(1, result.invalidRows());
    assertEquals(1, result.errors().getFirst().row());
  }

  @Test
  void rejectsValuesThatDoNotMatchTheConfiguredPattern() throws IOException {
    var result = new PatternImportService().importFromExcel(workbookWithText("invalid"));
    assertEquals(0, result.validRows());
    assertEquals(1, result.invalidRows());
  }

  @Test
  void autoConversionKeepsNumericValuesNumericForObjectFields() throws IOException {
    var result = new AutoImportService().importFromExcel(workbookWithText("12.50"));
    assertEquals(1, result.validRows());
    assertTrue(result.data().getFirst().value instanceof BigDecimal);
    assertEquals(0, new BigDecimal("12.50").compareTo((BigDecimal) result.data().getFirst().value));
  }

  @Test
  void rejectsDataWhenAnnotationHeaderDoesNotMatch() throws IOException {
    var result = new HeaderImportService().importFromExcel(workbookWithText("A-01"));
    assertEquals(0, result.totalRows());
    assertEquals(0, result.validRows());
    assertEquals(0, result.invalidRows());
    assertEquals(1, result.errors().size());
    assertEquals(ExcelImportProblemCode.HEADER_MISMATCH, result.errors().getFirst().problemCode());
    assertEquals(0, result.errors().getFirst().row());
  }

  @Test
  void rejectsInvalidWorksheetLayoutBeforeParsing() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ImportService()
                .importFromExcel(
                    new byte[0], new com.ravzeno.excel.model.import_.ExcelImportRequest(1, 2, 2)));
  }

  @Test
  void streamsTenThousandRowsForExportAndImport() throws IOException {
    Iterable<ExportRow> rows =
        () ->
            new Iterator<>() {
              private int current;

              @Override
              public boolean hasNext() {
                return current < 10_001;
              }

              @Override
              public ExportRow next() {
                current++;
                return new ExportRow("ROW-" + current, BigDecimal.ONE, LocalDate.of(2026, 10, 21));
              }
            };

    ByteArrayOutputStream output = new ByteArrayOutputStream();
    new ExportService().writeTo(rows, output);
    var imported = new ImportService().importFromExcel(output.toByteArray());

    assertEquals(10_001, imported.totalRows());
    assertEquals(10_001, imported.validRows());
  }

  private byte[] workbookWithFormula() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet();
      sheet.createRow(0).createCell(0).setCellValue("Amount");
      sheet.createRow(1).createCell(0).setCellFormula("2+3");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.write(output);
      return output.toByteArray();
    }
  }

  private byte[] workbookWithBlankCell() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet();
      sheet.createRow(0).createCell(0).setCellValue("Amount");
      sheet.getRow(0).createCell(1).setCellValue("Description");
      sheet.createRow(1).createCell(0).setBlank();
      sheet.getRow(1).createCell(1).setCellValue("row is not empty");
      workbook.write(output);
      return output.toByteArray();
    }
  }

  private byte[] workbookWithText(String value) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet();
      sheet.createRow(0).createCell(0).setCellValue("Code");
      sheet.createRow(1).createCell(0).setCellValue(value);
      workbook.write(output);
      return output.toByteArray();
    }
  }

  static class ExportRow {
    @ExcelColumn(
        name = "Code",
        index = 5,
        align = HorizontalAlignment.LEFT,
        allowedValues = {"A-01", "A-02"})
    String code;

    @ExcelColumn(name = "Amount", index = 10)
    BigDecimal amount;

    @ExcelColumn(name = "Date", index = 20)
    LocalDate date;

    ExportRow(String code, BigDecimal amount, LocalDate date) {
      this.code = code;
      this.amount = amount;
      this.date = date;
    }
  }

  static class ImportRow {
    @ExcelImportColumn(index = 0, required = true)
    String code;

    @ExcelImportColumn(index = 1, dataType = ExcelDataType.NUMERIC)
    BigDecimal amount;

    @ExcelImportColumn(index = 2, dataType = ExcelDataType.DATE)
    LocalDate date;
  }

  static class ExportService extends AbstractExcelExportService<ExportRow> {
    protected String getSheetName() {
      return "Orders";
    }

    protected Object[] mapToRow(ExportRow row) {
      return new Object[] {row.code, row.amount, row.date};
    }

    protected Class<ExportRow> getEntityClass() {
      return ExportRow.class;
    }
  }

  static class ImportService extends AbstractExcelImportService<ImportRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<ImportRow> getEntityClass() {
      return ImportRow.class;
    }
  }

  static class FormulaRow {
    @ExcelImportColumn(index = 0, dataType = ExcelDataType.NUMERIC)
    BigDecimal amount;
  }

  static class FormulaImportService extends AbstractExcelImportService<FormulaRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<FormulaRow> getEntityClass() {
      return FormulaRow.class;
    }
  }

  static class PrimitiveRow {
    @ExcelImportColumn(index = 0, required = true)
    int amount;
  }

  static class PrimitiveImportService extends AbstractExcelImportService<PrimitiveRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<PrimitiveRow> getEntityClass() {
      return PrimitiveRow.class;
    }
  }

  static class PatternRow {
    @ExcelImportColumn(index = 0, pattern = "[A-Z]{2}-\\d{2}")
    String code;
  }

  static class PatternImportService extends AbstractExcelImportService<PatternRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<PatternRow> getEntityClass() {
      return PatternRow.class;
    }
  }

  static class AutoRow {
    @ExcelImportColumn(index = 0)
    Object value;
  }

  static class AutoImportService extends AbstractExcelImportService<AutoRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<AutoRow> getEntityClass() {
      return AutoRow.class;
    }
  }

  static class HeaderRow {
    @ExcelImportColumn(index = 0, headerName = "Expected code")
    String code;
  }

  static class HeaderImportService extends AbstractExcelImportService<HeaderRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<HeaderRow> getEntityClass() {
      return HeaderRow.class;
    }
  }

  static class DateTimeRow {
    @ExcelImportColumn(index = 0, dataType = ExcelDataType.DATE)
    LocalDateTime value;
  }

  static class DateTimeImportService extends AbstractExcelImportService<DateTimeRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<DateTimeRow> getEntityClass() {
      return DateTimeRow.class;
    }
  }

  static class ConstrainedRow {
    @ExcelImportColumn(index = 0, dataType = ExcelDataType.NUMERIC, minValue = 10)
    BigDecimal value;
  }

  static class ConstrainedImportService extends AbstractExcelImportService<ConstrainedRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<ConstrainedRow> getEntityClass() {
      return ConstrainedRow.class;
    }
  }
}
