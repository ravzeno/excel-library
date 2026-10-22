package com.ravzeno.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.ravzeno.excel.annotation.ExcelColumn;
import com.ravzeno.excel.annotation.ExcelImportColumn;
import com.ravzeno.excel.base.AbstractExcelExportService;
import com.ravzeno.excel.base.AbstractExcelImportService;
import com.ravzeno.excel.enums.ExcelDataType;
import com.ravzeno.excel.model.export.ExcelExportResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class ExcelLibraryTest {
    @Test
    void exportsAndImportsAnnotatedRows() throws IOException {
        ExcelExportResult exported = new ExportService().export(List.of(new ExportRow("A-01", new BigDecimal("12.50"), LocalDate.of(2026, 10, 21))));
        assertTrue(exported.bytes().length > 0);
        var result = new ImportService().importFromExcel(exported.bytes());
        assertEquals(1, result.validRows());
        assertEquals("A-01", result.data().getFirst().code);
        assertEquals(0, new BigDecimal("12.50").compareTo(result.data().getFirst().amount));
        assertEquals(LocalDate.of(2026, 10, 21), result.data().getFirst().date);
    }

    @Test
    void exportsStylesForNonContiguousAnnotationIndexes() throws IOException {
        ExcelExportResult exported = new ExportService().export(List.of(new ExportRow("A-01", new BigDecimal("12.50"), LocalDate.of(2026, 10, 21))));
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(exported.bytes()))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("Code", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Amount", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(HorizontalAlignment.LEFT, sheet.getRow(1).getCell(0).getCellStyle().getAlignment());
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
    void streamsTenThousandRowsForExportAndImport() throws IOException {
        Iterable<ExportRow> rows = () -> new Iterator<>() {
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
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet();
            sheet.createRow(0).createCell(0).setCellValue("Amount");
            sheet.createRow(1).createCell(0).setCellFormula("2+3");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            workbook.write(output);
            return output.toByteArray();
        }
    }

    private byte[] workbookWithBlankCell() throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
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
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet();
            sheet.createRow(0).createCell(0).setCellValue("Code");
            sheet.createRow(1).createCell(0).setCellValue(value);
            workbook.write(output);
            return output.toByteArray();
        }
    }

    static class ExportRow {
        @ExcelColumn(name = "Code", index = 5, align = HorizontalAlignment.LEFT) String code;
        @ExcelColumn(name = "Amount", index = 10) BigDecimal amount;
        @ExcelColumn(name = "Date", index = 20) LocalDate date;
        ExportRow(String code, BigDecimal amount, LocalDate date) { this.code = code; this.amount = amount; this.date = date; }
    }
    static class ImportRow {
        @ExcelImportColumn(index = 0, required = true) String code;
        @ExcelImportColumn(index = 1, dataType = ExcelDataType.NUMERIC) BigDecimal amount;
        @ExcelImportColumn(index = 2, dataType = ExcelDataType.DATE) LocalDate date;
    }
    static class ExportService extends AbstractExcelExportService<ExportRow> {
        protected String getSheetName() { return "Orders"; }
        protected Object[] mapToRow(ExportRow row) { return new Object[] {row.code, row.amount, row.date}; }
        protected Class<ExportRow> getEntityClass() { return ExportRow.class; }
    }
    static class ImportService extends AbstractExcelImportService<ImportRow> {
        protected int getSheetNo() { return 1; }
        protected int getHeaderRow() { return 1; }
        protected int getDataRow() { return 2; }
        protected Class<ImportRow> getEntityClass() { return ImportRow.class; }
    }
    static class FormulaRow { @ExcelImportColumn(index = 0, dataType = ExcelDataType.NUMERIC) BigDecimal amount; }
    static class FormulaImportService extends AbstractExcelImportService<FormulaRow> {
        protected int getSheetNo() { return 1; }
        protected int getHeaderRow() { return 1; }
        protected int getDataRow() { return 2; }
        protected Class<FormulaRow> getEntityClass() { return FormulaRow.class; }
    }
    static class PrimitiveRow { @ExcelImportColumn(index = 0, required = true) int amount; }
    static class PrimitiveImportService extends AbstractExcelImportService<PrimitiveRow> {
        protected int getSheetNo() { return 1; }
        protected int getHeaderRow() { return 1; }
        protected int getDataRow() { return 2; }
        protected Class<PrimitiveRow> getEntityClass() { return PrimitiveRow.class; }
    }
    static class PatternRow { @ExcelImportColumn(index = 0, pattern = "[A-Z]{2}-\\d{2}") String code; }
    static class PatternImportService extends AbstractExcelImportService<PatternRow> {
        protected int getSheetNo() { return 1; }
        protected int getHeaderRow() { return 1; }
        protected int getDataRow() { return 2; }
        protected Class<PatternRow> getEntityClass() { return PatternRow.class; }
    }
    static class AutoRow { @ExcelImportColumn(index = 0) Object value; }
    static class AutoImportService extends AbstractExcelImportService<AutoRow> {
        protected int getSheetNo() { return 1; }
        protected int getHeaderRow() { return 1; }
        protected int getDataRow() { return 2; }
        protected Class<AutoRow> getEntityClass() { return AutoRow.class; }
    }
}
