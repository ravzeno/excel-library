package com.ravzeno.excel;

import static org.junit.jupiter.api.Assertions.*;

import com.ravzeno.excel.annotation.ExcelColumn;
import com.ravzeno.excel.annotation.ExcelImportColumn;
import com.ravzeno.excel.base.AbstractExcelExportService;
import com.ravzeno.excel.base.AbstractExcelImportService;
import com.ravzeno.excel.enums.ExcelDataType;
import com.ravzeno.excel.enums.ExcelDateFormat;
import com.ravzeno.excel.enums.ExcelImportProblemCode;
import com.ravzeno.excel.enums.ExcelNumberFormat;
import com.ravzeno.excel.model.export.ExcelColumnDef;
import com.ravzeno.excel.model.export.ExcelExportRequest;
import com.ravzeno.excel.model.export.ExcelExportResult;
import com.ravzeno.excel.model.export.ExcelHeaderColumn;
import com.ravzeno.excel.model.import_.ExcelImportError;
import com.ravzeno.excel.model.import_.ExcelImportErrorWorkbook;
import com.ravzeno.excel.model.import_.ExcelImportRequest;
import com.ravzeno.excel.util.ExcelUtils;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

class ExcelPublicApiTest {
  @Nested
  class ValueAndUtilityTests {
    @Test
    void parsesSupportedDatesNumbersAndDateShapes() {
      assertEquals(
          LocalDate.of(2026, 10, 21), ExcelUtils.parseDate("2026-10-21", ExcelDateFormat.ISO_DATE));
      assertEquals(
          LocalDate.of(2026, 10, 21),
          ExcelUtils.parseDate("21/10/2569", ExcelDateFormat.THAI_DATE));
      assertEquals(
          LocalDate.of(2026, 10, 21),
          ExcelUtils.parseDate("21-10-2026", ExcelDateFormat.ISO_DATE, "dd-MM-yyyy"));
      assertEquals(LocalDate.of(1900, 1, 1), ExcelUtils.parseDate("2.0", ExcelDateFormat.ISO_DATE));
      assertNull(ExcelUtils.parseDate("not a date", ExcelDateFormat.ISO_DATE));
      assertEquals(
          LocalDateTime.of(2026, 10, 21, 14, 30),
          ExcelUtils.parseDateTime(
              "21-10-2026 14:30", ExcelDateFormat.ISO_DATE, "dd-MM-yyyy HH:mm"));
      assertEquals(
          LocalDateTime.of(2026, 10, 21, 0, 0),
          ExcelUtils.parseDateTime("21-10-2026", ExcelDateFormat.ISO_DATE, "dd-MM-yyyy"));
      assertEquals(
          LocalDateTime.of(2026, 11, 1, 3, 4),
          ExcelUtils.parseDateTime("1/11/2026 3:04", ExcelDateFormat.THAI_DATETIME));
      assertNull(ExcelUtils.parseDateTime(null, ExcelDateFormat.ISO_DATE));
      assertTrue(ExcelUtils.isNumeric("1,234.50"));
      assertFalse(ExcelUtils.isNumeric("  "));
      assertTrue(ExcelUtils.isDate("21/10/2026"));
      assertFalse(ExcelUtils.isDate("tomorrow"));
    }

    @Test
    void normalizesImportRequestsAndExportPayloads() {
      assertEquals(1, new ExcelImportRequest(null, null, null).getSheetNoOrDefault());
      assertEquals(1, new ExcelImportRequest(null, null, null).getHeaderRowOrDefault());
      assertEquals(2, new ExcelImportRequest(null, null, null).getDataRowOrDefault());
      var request = new ExcelExportRequest<String>(null, " ", "", null, null);
      assertEquals(List.of(), request.data());
      assertEquals("Sheet1", request.sheetName());
      assertEquals("export.xlsx", request.filename());
      assertTrue(request.freezeHeader());
      var source = new ArrayList<>(List.of("before"));
      var copied = new ExcelExportRequest<>(source, "Rows", "rows.xlsx", null, null);
      source.add("after");
      assertEquals(List.of("before"), copied.data());
      assertThrows(
          NullPointerException.class,
          () ->
              new ExcelExportRequest<>(
                  java.util.Arrays.asList("x", null), "Rows", "rows.xlsx", null, null));
      assertEquals(
          ExcelExportResult.XLSX_CONTENT_TYPE,
          new ExcelExportResult(new byte[0], "a.xlsx", " ").contentType());
      assertEquals(
          "application/test",
          new ExcelExportResult(new byte[0], "a", "application/test").contentType());
    }

    @Test
    void resolvesColumnMetadataAndRejectsInvalidAnnotationColor() throws Exception {
      var field = StyledRow.class.getDeclaredField("amount");
      ExcelColumnDef definition = ExcelColumnDef.fromField(field);
      assertEquals("Amount", definition.name());
      assertEquals(2, definition.index());
      assertEquals(12 * ExcelColumnDef.WPC, definition.width());
      assertEquals("#,##0.00", definition.numberFormat().getPattern());
      assertArrayEquals(new String[] {"A", "B"}, definition.allowedValues());
      assertNull(ExcelColumnDef.fromField(StyledRow.class.getDeclaredField("plain")));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelColumnDef.fromField(InvalidColorRow.class.getDeclaredField("value")));
      assertEquals("#,##0.00", ExcelNumberFormat.DECIMAL_2.getPattern());
    }
  }

  @Nested
  class ImportTests {
    @Test
    void importsXlsxFromInputStreamUsingDefaultWorksheetLayout() throws IOException {
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var input =
          new TrackingInputStream(
              workbookWithRows(
                  new String[] {"Code", "Description"}, new String[][] {{"A-7", "streamed"}}));

      var result = importer.importFromExcel(input);

      assertEquals(1, result.validRows());
      assertEquals("A-7", result.data().getFirst().code);
      assertEquals("streamed", result.data().getFirst().description);
      assertFalse(input.closed);
    }

    @Test
    void importsCsvWithQuotedCommasAndStreamsValidRowsInBatches() throws IOException {
      String csv = "Code,Description\r\nA-1,\"a,b\"\r\nA-2,\"say \"\"hi\"\"\"\r\n";
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var result =
          importer.importFromCsv(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));
      assertEquals(2, result.validRows());
      assertEquals("a,b", result.data().getFirst().description);
      assertEquals("say \"hi\"", result.data().get(1).description);

      var batches = new ArrayList<List<CsvRow>>();
      var summary =
          importer.importInBatches(
              new ByteArrayInputStream(
                  workbookWithRows(
                      new String[] {"Code", "Description"},
                      new String[][] {{"A-1", "a,b"}, {"A-2", "say hi"}})),
              1,
              batches::add);
      assertEquals(2, summary.totalRows());
      assertEquals(2, summary.validRows());
      assertEquals(2, batches.size());
      assertEquals(1, batches.getFirst().size());
      batches.clear();
      importer.importInBatches(
          new ByteArrayInputStream(
              workbookWithRows(
                  new String[] {"Code", "Description"},
                  new String[][] {{"A-1", "a,b"}, {"A-2", "say hi"}, {"A-3", "tail"}})),
          2,
          batches::add);
      assertEquals(List.of(2, 1), batches.stream().map(List::size).toList());

      List<CsvRow> visited = new ArrayList<>();
      importer.importEach(
          new ByteArrayInputStream(
              workbookWithRows(
                  new String[] {"Code", "Description"},
                  new String[][] {{"A-1", "a,b"}, {"A-2", "hi"}})),
          visited::add);
      assertEquals(2, visited.size());
      assertThrows(
          IllegalArgumentException.class,
          () -> importer.importInBatches(new ByteArrayInputStream(new byte[0]), 0, batches::add));
    }

    @Test
    void importsCsvUsingAnExplicitCharsetAndReportsMalformedCsv() throws IOException {
      String csv = "Code,Description\nA-1,café\n";
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var result =
          importer.importFromCsv(
              new ByteArrayInputStream(csv.getBytes(StandardCharsets.ISO_8859_1)),
              StandardCharsets.ISO_8859_1,
              null);
      assertEquals("café", result.data().getFirst().description);
      assertThrows(
          IOException.class,
          () ->
              importer.importFromCsv(
                  new ByteArrayInputStream(
                      "Code,Description\nA-1,\"unfinished".getBytes(StandardCharsets.UTF_8))));
      assertThrows(NullPointerException.class, () -> importer.importFromCsv(null));
    }

    @Test
    void preservesCsvInputOwnershipAndAcceptsUtf8Bom() throws IOException {
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var input =
          new TrackingInputStream(
              "\uFEFFCode,Description\nA-1,with BOM\n".getBytes(StandardCharsets.UTF_8));

      var result = importer.importFromCsv(input);

      assertFalse(input.closed);
      assertEquals(1, result.validRows());
      assertEquals("A-1", result.data().getFirst().code);
    }

    @Test
    void appliesCustomDatePatternsAndConfiguredInstantZone() throws IOException {
      var importer = new ZonedDateImportService();
      var result =
          importer.importFromExcel(
              workbookWithRows(new String[] {"Created"}, new String[][] {{"21-10-2026 14:30"}}));
      assertEquals(Instant.parse("2026-10-21T07:30:00Z"), result.data().getFirst().createdAt);
    }

    @Test
    void importsThaiDatesWithSingleDigitDayAndMonth() throws IOException {
      var importer = AbstractExcelImportService.of(VariableWidthDateRow.class, 1, 1, 2);
      var result =
          importer.importFromExcel(
              workbookWithRows(new String[] {"Date"}, new String[][] {{"1/11/2026"}}));
      assertTrue(result.errors().isEmpty(), "Expected 1/11/2026 to be accepted as day/month/year");
      assertEquals(LocalDate.of(2026, 11, 1), result.data().getFirst().date);
    }

    @Test
    void usesExplicitDatePatternForObjectFieldsInAutoMode() throws IOException {
      var importer = AbstractExcelImportService.of(ObjectDateRow.class, 1, 1, 2);
      var result =
          importer.importFromExcel(
              workbookWithRows(new String[] {"Date"}, new String[][] {{"20261101"}}));

      assertEquals(LocalDate.of(2026, 11, 1), result.data().getFirst().value);
    }

    @Test
    void streamsOnlyValidRowsAndReturnsValidationErrors() throws IOException {
      var importer = AbstractExcelImportService.of(RequiredCsvRow.class, 1, 1, 2);
      byte[] workbook =
          workbookWithRows(
              new String[] {"Code", "Note"},
              new String[][] {{"OK", "one"}, {"", "required value missing"}, {"YES", "three"}});
      List<RequiredCsvRow> rows = new ArrayList<>();
      var summary = importer.importEach(new ByteArrayInputStream(workbook), rows::add);
      assertEquals(3, summary.totalRows());
      assertEquals(2, summary.validRows());
      assertEquals(1, summary.invalidRows());
      assertEquals(1, summary.errors().size());
      assertEquals(2, summary.errors().getFirst().row());
      assertEquals(2, rows.size());
      assertThrows(
          NullPointerException.class,
          () ->
              importer.importEach(
                  new ByteArrayInputStream(workbook),
                  (java.util.function.Consumer<RequiredCsvRow>) null));
    }

    @Test
    void propagatesExceptionsThrownByRowConsumer() throws IOException {
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var expected = new IllegalStateException("consumer failed");
      var actual =
          assertThrows(
              IllegalStateException.class,
              () ->
                  importer.importEach(
                      new ByteArrayInputStream(
                          workbookWithRows(
                              new String[] {"Code", "Description"},
                              new String[][] {{"A-1", "one"}})),
                      row -> {
                        throw expected;
                      }));

      assertSame(expected, actual);
    }

    @Test
    void validatesPerCallWorksheetLayoutAndMissingHeaders() throws IOException {
      byte[] workbook = workbookWithRows(new String[] {"Ignored"}, new String[][] {{"skip"}});
      var importer = AbstractExcelImportService.of(CsvRow.class, 1, 1, 2);
      var missing = importer.importFromExcel(workbookWithRows(new String[0], new String[0][]));
      assertEquals(
          ExcelImportProblemCode.HEADER_ROW_NOT_FOUND, missing.errors().getFirst().problemCode());
      assertThrows(
          IllegalArgumentException.class,
          () -> importer.importFromExcel(workbook, new ExcelImportRequest(2, 1, 2)));
      assertThrows(
          IllegalArgumentException.class,
          () -> importer.importFromExcel(workbook, new ExcelImportRequest(1, 0, 1)));
      var validRequestResult =
          importer.importFromExcel(
              new ByteArrayInputStream(
                  workbookWithRows(
                      new String[] {"Code", "Description"}, new String[][] {{"A", "one"}})),
              new ExcelImportRequest(1, 1, 2));
      assertEquals(1, validRequestResult.validRows());
    }

    @Test
    void appliesLengthRangeAllowedValueAndPatternValidations() throws IOException {
      var importer = AbstractExcelImportService.of(ValidationRow.class, 1, 1, 2);
      var result =
          importer.importFromExcel(
              workbookWithRows(
                  new String[] {"Code", "Category", "Count", "Pattern"},
                  new String[][] {
                    {"x", "A", "2", "AB-12"},
                    {"valid", "Z", "2", "AB-12"},
                    {"valid", "A", "9", "AB-12"},
                    {"valid", "A", "2", "bad"}
                  }));
      assertEquals(0, result.validRows());
      assertEquals(4, result.invalidRows());
      assertEquals(
          List.of(
              ExcelImportProblemCode.MIN_LENGTH,
              ExcelImportProblemCode.VALUE_NOT_ALLOWED,
              ExcelImportProblemCode.MAX_VALUE,
              ExcelImportProblemCode.PATTERN_MISMATCH),
          result.errors().stream().map(ExcelImportError::problemCode).toList());
    }

    @Test
    void reportsMalformedConfiguredPattern() throws IOException {
      var importer = AbstractExcelImportService.of(InvalidPatternRow.class, 1, 1, 2);
      var result =
          importer.importFromExcel(
              workbookWithRows(new String[] {"Code"}, new String[][] {{"anything"}}));
      assertEquals(
          ExcelImportProblemCode.INVALID_PATTERN_REGEX, result.errors().getFirst().problemCode());
    }

    @Test
    void appliesSubclassHeaderAndCrossFieldRules() throws IOException {
      var importer = new CustomValidationImporter();
      var result =
          importer.importFromExcel(
              workbookWithRows(
                  new String[] {"From", "To"}, new String[][] {{"reject", "ok"}, {"ok", "ok"}}));
      assertEquals(1, result.validRows());
      assertEquals(1, result.invalidRows());
      assertEquals(ExcelImportProblemCode.MAPPING_ERROR, result.errors().getFirst().problemCode());
      var badHeaders =
          importer.importFromExcel(
              workbookWithRows(new String[] {"Wrong", "To"}, new String[][] {{"ok", "ok"}}));
      assertEquals(
          ExcelImportProblemCode.HEADER_MISMATCH, badHeaders.errors().getFirst().problemCode());
    }

    @Test
    void callsHeaderValidationHookWithTheUploadedHeaderRow() throws IOException {
      var importer = new HeaderHookImporter();
      var result =
          importer.importFromExcel(
              workbookWithRows(
                  new String[] {"Code", "Description"}, new String[][] {{"A-1", "one"}}));
      assertEquals(1, result.validRows());
      assertEquals(List.of("Code", "Description"), importer.seenHeaders);
    }
  }

  @Nested
  class ExportTests {
    @Test
    void exportsAnnotatedValuesWithoutASeparateRowMapper() throws IOException {
      var exporter = AbstractExcelExportService.of(ExportValues.class, "Direct");
      var result =
          exporter.export(List.of(new ExportValues("direct", 7, true, LocalDate.of(2026, 11, 1))));
      try (var workbook = new XSSFWorkbook(new ByteArrayInputStream(result.bytes()))) {
        var row = workbook.getSheet("Direct").getRow(1);
        assertEquals("direct", row.getCell(0).getStringCellValue());
        assertEquals(7d, row.getCell(1).getNumericCellValue());
        assertTrue(row.getCell(2).getBooleanCellValue());
        assertEquals(
            LocalDate.of(2026, 11, 1), row.getCell(3).getLocalDateTimeCellValue().toLocalDate());
      }
    }

    @Test
    void buildsSheetsWithManualStylesAndRejectsInvalidRows() throws IOException {
      try (var workbook = new XSSFWorkbook()) {
        var headerColor = new XSSFColor(new byte[] {(byte) 0x11, (byte) 0x22, (byte) 0x33}, null);
        var builder =
            new ExcelDataBuilder(
                workbook,
                "Manual",
                new ExcelHeaderColumn[] {
                  ExcelHeaderColumn.create("Value", 20 * ExcelHeaderColumn.WPC, headerColor)
                },
                List.<Object[]>of(new Object[] {"sample"}),
                new ExcelDataBuilderStyle(null, null, 24f, "Arial"));
        var sheet = builder.createSheet();
        assertEquals("Value", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals(24f, sheet.getRow(0).getHeightInPoints());
        assertEquals("sample", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals(20 * ExcelHeaderColumn.WPC, sheet.getColumnWidth(0));
        assertEquals(
            HorizontalAlignment.CENTER, sheet.getRow(1).getCell(0).getCellStyle().getAlignment());
      }
      assertThrows(
          IllegalArgumentException.class,
          () ->
              new ExcelDataBuilder(
                  new XSSFWorkbook(), " ", new ExcelHeaderColumn[0], List.of(), null));
      try (var workbook = new XSSFWorkbook()) {
        var tooMany =
            new ExcelDataBuilder(
                workbook,
                "Rows",
                new ExcelHeaderColumn[] {ExcelHeaderColumn.create("Only", 1000, null)},
                List.<Object[]>of(new Object[] {"one", "two"}),
                null);
        assertThrows(IllegalArgumentException.class, tooMany::createSheet);
      }
    }

    @Test
    void writesConfiguredExportRequestsToCallerOwnedStreams() throws IOException {
      var exporter =
          AbstractExcelExportService.of(
              ExportValues.class,
              "Default",
              row -> new Object[] {row.text, row.count, row.enabled, row.date, null});
      var request =
          new ExcelExportRequest<>(
              List.of(new ExportValues("text", 3, true, LocalDate.of(2026, 10, 21))),
              "Configured",
              "custom.xlsx",
              30f,
              "Arial",
              false);
      var output = new ByteArrayOutputStream();
      exporter.writeTo(request, output);
      try (var workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
        var sheet = workbook.getSheet("Configured");
        assertNotNull(sheet);
        assertEquals("custom.xlsx", exporter.export(request).filename());
        assertEquals(30f, sheet.getRow(0).getHeightInPoints());
        assertEquals("text", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals(3d, sheet.getRow(1).getCell(1).getNumericCellValue());
        assertTrue(sheet.getRow(1).getCell(2).getBooleanCellValue());
        assertEquals(
            LocalDate.of(2026, 10, 21),
            sheet.getRow(1).getCell(3).getLocalDateTimeCellValue().toLocalDate());
        assertEquals(
            org.apache.poi.ss.usermodel.CellType.BLANK, sheet.getRow(1).getCell(4).getCellType());
        assertNull(sheet.getPaneInformation());
      }
    }

    @Test
    void rejectsInvalidExportSchemasRequestsAndRows() {
      assertThrows(
          NullPointerException.class,
          () -> AbstractExcelExportService.of(null, "x", x -> new Object[0]));
      assertThrows(
          NullPointerException.class,
          () -> AbstractExcelExportService.of(ExportValues.class, null, x -> new Object[0]));
      assertThrows(
          NullPointerException.class,
          () -> AbstractExcelExportService.of(ExportValues.class, "x", null));
      var noColumns = AbstractExcelExportService.of(NoExportColumns.class, "x", x -> new Object[0]);
      assertThrows(
          IllegalArgumentException.class, () -> noColumns.export(List.of(new NoExportColumns())));
      var duplicates =
          AbstractExcelExportService.of(
              DuplicateExportColumns.class, "x", x -> new Object[] {"a", "b"});
      assertThrows(
          IllegalArgumentException.class,
          () -> duplicates.export(List.of(new DuplicateExportColumns())));
      var exporter =
          AbstractExcelExportService.of(
              ExportValues.class,
              "x",
              x -> new Object[] {x.text, x.count, x.enabled, x.date, null, "extra"});
      assertThrows(
          IllegalArgumentException.class,
          () -> exporter.export(List.of(new ExportValues("x", 1, true, LocalDate.now()))));
      assertThrows(NullPointerException.class, () -> exporter.export((List<ExportValues>) null));
      assertThrows(
          IllegalArgumentException.class, () -> exporter.exportWorkbook(List.of(), "empty.xlsx"));
    }

    @Test
    void errorWorkbookHighlightsCellsAndSkipsRowLevelErrors() throws IOException {
      byte[] source = workbookWithRows(new String[] {"A"}, new String[][] {{"bad"}});
      var errors =
          List.of(
              new ExcelImportError(
                  1, 0, "A", "bad", ExcelImportProblemCode.INVALID_FORMAT, "bad value"),
              new ExcelImportError(
                  1, -1, null, null, ExcelImportProblemCode.MAPPING_ERROR, "row error"));
      byte[] marked = ExcelImportErrorWorkbook.create(source, errors);
      try (var workbook = new XSSFWorkbook(new ByteArrayInputStream(marked))) {
        var cell = workbook.getSheetAt(0).getRow(1).getCell(0);
        assertEquals("bad value", cell.getCellComment().getString().getString());
        assertEquals(FillPatternType.SOLID_FOREGROUND, cell.getCellStyle().getFillPattern());
        assertEquals(IndexedColors.ROSE.getIndex(), cell.getCellStyle().getFillForegroundColor());
      }
      assertThrows(NullPointerException.class, () -> ExcelImportErrorWorkbook.create(null, errors));
      assertThrows(
          NullPointerException.class,
          () -> ExcelImportErrorWorkbook.create(source, (List<ExcelImportError>) null));
      assertThrows(
          IOException.class, () -> ExcelImportErrorWorkbook.create(new byte[] {1}, List.of()));
    }

    @Test
    void exposesHeaderColumnFactoryAndExportResultDefaults() {
      var header = ExcelHeaderColumn.create("Name", 12 * ExcelHeaderColumn.WPC, null);
      assertEquals("Name", header.name());
      assertEquals(12 * 256, header.width());
      assertEquals(
          ExcelExportResult.XLSX_CONTENT_TYPE,
          new ExcelExportResult(new byte[] {1}, "x.xlsx").contentType());
    }
  }

  private static byte[] workbookWithRows(String[] headers, String[][] values) throws IOException {
    try (var workbook = new XSSFWorkbook();
        var bytes = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("Input");
      if (headers.length > 0) {
        var header = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) header.createCell(i).setCellValue(headers[i]);
      }
      for (int r = 0; r < values.length; r++) {
        var row = sheet.createRow(r + 1);
        for (int c = 0; c < values[r].length; c++) row.createCell(c).setCellValue(values[r][c]);
      }
      workbook.write(bytes);
      return bytes.toByteArray();
    }
  }

  private static class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    private TrackingInputStream(byte[] bytes) {
      super(bytes);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }
  }

  static class CsvRow {
    @ExcelImportColumn(index = 0, headerName = "Code", required = true)
    String code;

    @ExcelImportColumn(index = 1, headerName = "Description")
    String description;
  }

  static class RequiredCsvRow {
    @ExcelImportColumn(index = 0, required = true)
    String code;

    @ExcelImportColumn(index = 1)
    String note;
  }

  static class ValidationRow {
    @ExcelImportColumn(index = 0, minLength = 3, maxLength = 8)
    String code;

    @ExcelImportColumn(
        index = 1,
        allowedValues = {"A", "B"})
    String category;

    @ExcelImportColumn(index = 2, dataType = ExcelDataType.NUMERIC, minValue = 1, maxValue = 5)
    Integer count;

    @ExcelImportColumn(index = 3, pattern = "[A-Z]{2}-\\d{2}")
    String pattern;
  }

  static class InvalidPatternRow {
    @ExcelImportColumn(index = 0, pattern = "[")
    String code;
  }

  static class CustomValidationRow {
    @ExcelImportColumn(index = 0, headerName = "From")
    String from;

    @ExcelImportColumn(index = 1, headerName = "To")
    String to;
  }

  static class CustomValidationImporter extends AbstractExcelImportService<CustomValidationRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<CustomValidationRow> getEntityClass() {
      return CustomValidationRow.class;
    }

    protected List<ExcelImportError> validateRow(CustomValidationRow row, int worksheetRow) {
      return row.from.equals("reject")
          ? List.of(
              new ExcelImportError(
                  worksheetRow,
                  -1,
                  null,
                  row.from,
                  ExcelImportProblemCode.MAPPING_ERROR,
                  "row rejected"))
          : List.of();
    }
  }

  static class HeaderHookImporter extends AbstractExcelImportService<CsvRow> {
    List<String> seenHeaders = List.of();

    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<CsvRow> getEntityClass() {
      return CsvRow.class;
    }

    protected void validateHeaders(List<String> headers) {
      seenHeaders = List.copyOf(headers);
    }
  }

  static class ZonedDateRow {
    @ExcelImportColumn(index = 0, dataType = ExcelDataType.DATE, datePattern = "dd-MM-yyyy HH:mm")
    Instant createdAt;
  }

  static class VariableWidthDateRow {
    @ExcelImportColumn(index = 0, dataType = ExcelDataType.DATE)
    LocalDate date;
  }

  static class ObjectDateRow {
    @ExcelImportColumn(index = 0, datePattern = "yyyyMMdd")
    Object value;
  }

  static class ZonedDateImportService extends AbstractExcelImportService<ZonedDateRow> {
    protected int getSheetNo() {
      return 1;
    }

    protected int getHeaderRow() {
      return 1;
    }

    protected int getDataRow() {
      return 2;
    }

    protected Class<ZonedDateRow> getEntityClass() {
      return ZonedDateRow.class;
    }

    protected ZoneId getZoneId() {
      return ZoneId.of("Asia/Bangkok");
    }
  }

  static class StyledRow {
    @ExcelColumn(
        name = "Amount",
        index = 2,
        width = 12,
        numberFormat = ExcelNumberFormat.DECIMAL_2,
        allowedValues = {"A", "B"})
    BigDecimal amount;

    String plain;
  }

  static class InvalidColorRow {
    @ExcelColumn(name = "Invalid", index = 0, headerCellColor = "blue")
    String value;
  }

  static class ExportValues {
    @ExcelColumn(name = "Text", index = 0)
    String text;

    @ExcelColumn(name = "Count", index = 1)
    int count;

    @ExcelColumn(name = "Enabled", index = 2)
    boolean enabled;

    @ExcelColumn(name = "Date", index = 3)
    LocalDate date;

    @ExcelColumn(name = "Blank", index = 4)
    String blank;

    ExportValues(String text, int count, boolean enabled, LocalDate date) {
      this.text = text;
      this.count = count;
      this.enabled = enabled;
      this.date = date;
    }
  }

  static class NoExportColumns {}

  static class DuplicateExportColumns {
    @ExcelColumn(name = "A", index = 0)
    String first;

    @ExcelColumn(name = "B", index = 0)
    String second;
  }
}
