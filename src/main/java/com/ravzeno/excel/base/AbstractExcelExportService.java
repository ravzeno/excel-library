package com.ravzeno.excel.base;

import com.ravzeno.excel.ExcelDataBuilder;
import com.ravzeno.excel.ExcelDataBuilderStyle;
import com.ravzeno.excel.annotation.ExcelColumn;
import com.ravzeno.excel.model.export.ExcelColumnDef;
import com.ravzeno.excel.model.export.ExcelExportRequest;
import com.ravzeno.excel.model.export.ExcelExportResult;
import com.ravzeno.excel.model.export.ExcelHeaderColumn;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * คลาสพื้นฐานสำหรับ export ไฟล์ XLSX โดยกำหนด schema จาก annotation
 *
 * <p>คลาสลูกกำหนด schema ผ่าน field ที่ติด {@code @ExcelColumn} และแปลงข้อมูลแต่ละรายการเป็น
 * หนึ่งแถว บริการนี้เขียนข้อมูลแบบ lazy จึงใช้ {@link Iterable} ที่ดึงข้อมูลเป็นหน้า ๆ
 * จากฐานข้อมูลได้ โดยไม่ต้องเก็บทุกแถวไว้ในหน่วยความจำ
 *
 * @param <T> ชนิดข้อมูล domain ที่แทนหนึ่งแถวใน spreadsheet
 */
public abstract class AbstractExcelExportService<T> {
  /** constructor สำหรับคลาสลูก */
  protected AbstractExcelExportService() {}

  /**
   * Creates an annotation-based exporter without requiring a dedicated subclass.
   *
   * @param entityClass type that declares {@code @ExcelColumn} fields
   * @param sheetName visible worksheet name
   * @param rowMapper converts each source item to cells in annotation order
   * @param <T> source row type
   * @return a ready-to-use exporter
   */
  public static <T> AbstractExcelExportService<T> of(
      Class<T> entityClass, String sheetName, Function<? super T, Object[]> rowMapper) {
    Objects.requireNonNull(entityClass, "entityClass must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(rowMapper, "rowMapper must not be null");
    return new AbstractExcelExportService<>() {
      @Override
      protected String getSheetName() {
        return sheetName;
      }

      @Override
      protected Object[] mapToRow(T value) {
        return rowMapper.apply(value);
      }

      @Override
      protected Class<T> getEntityClass() {
        return entityClass;
      }
    };
  }

  /**
   * ระบุชื่อ worksheet เริ่มต้นที่จะแสดงในไฟล์
   *
   * @return ชื่อ worksheet
   */
  protected abstract String getSheetName();

  /**
   * แปลงข้อมูล domain หนึ่งรายการเป็นค่า cell ตามลำดับเดียวกับ {@code @ExcelColumn}
   *
   * @param value ข้อมูล domain ต้นทาง
   * @return array ค่า cell ที่ต้องไม่เป็น {@code null}
   */
  protected abstract Object[] mapToRow(T value);

  /**
   * ระบุ type ที่ประกาศ annotation {@code @ExcelColumn} สำหรับ schema
   *
   * @return type ของ schema ที่ติด annotation
   */
  protected abstract Class<T> getEntityClass();

  /**
   * Export list ที่อยู่ในหน่วยความจำ โดยรักษาลำดับเดิมของ list
   *
   * @param data แถวข้อมูลที่ต้องการ export
   * @return เนื้อหา XLSX ที่อยู่ในหน่วยความจำ
   * @throws IOException เมื่อ POI เขียนเอกสาร XLSX ไม่สำเร็จ
   */
  public ExcelExportResult export(List<T> data) throws IOException {
    return export((Iterable<T>) data);
  }

  /**
   * Export {@link Iterable} แบบ lazy และคืนผลลัพธ์เป็น byte ที่อยู่ในหน่วยความจำ
   *
   * <p>แม้แถวจะถูกสร้างแบบ streaming แต่ byte array ของ XLSX ที่เสร็จแล้วจะยังอยู่ในหน่วยความจำ
   * สำหรับ HTTP response หรือไฟล์ขนาดใหญ่มาก ให้ใช้ {@link #writeTo(Iterable, OutputStream)} แทน
   *
   * @param data แถวต้นทางแบบ lazy หรือในหน่วยความจำ
   * @return เนื้อหา XLSX ที่อยู่ในหน่วยความจำ
   * @throws IOException เมื่อ POI เขียนเอกสาร XLSX ไม่สำเร็จ
   */
  public ExcelExportResult export(Iterable<T> data) throws IOException {
    return exportRows(
        Objects.requireNonNull(data, "data must not be null"),
        new ExcelExportRequest<>(List.of(), getSheetName(), "export.xlsx", null, null));
  }

  /**
   * Export แถวและตัวเลือกที่อยู่ใน request
   *
   * @param request แถวข้อมูลและการตั้งค่ารูปแบบ export
   * @return เนื้อหา XLSX ที่อยู่ในหน่วยความจำ
   * @throws IOException เมื่อ POI เขียนเอกสาร XLSX ไม่สำเร็จ
   */
  public ExcelExportResult export(ExcelExportRequest<T> request) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return exportRows(request.data(), request);
  }

  /**
   * เขียนแถวที่สร้างแบบ lazy ลง output stream โดยตรง โดยไม่เก็บ byte ของ XLSX ไว้
   * ผู้เรียกเป็นเจ้าของและต้องปิด {@code output} เอง
   *
   * @param data แถวต้นทางแบบ lazy หรือในหน่วยความจำ
   * @param output stream ปลายทาง ซึ่งจะยังเปิดอยู่หลังเขียนเสร็จ
   * @throws IOException เมื่อ POI เขียนเอกสาร XLSX ไม่สำเร็จ
   */
  public void writeTo(Iterable<T> data, OutputStream output) throws IOException {
    writeRows(
        Objects.requireNonNull(data, "data must not be null"),
        new ExcelExportRequest<>(List.of(), getSheetName(), "export.xlsx", null, null),
        output);
  }

  /**
   * เขียนเอกสาร XLSX ลง output stream โดยตรงโดยใช้ตัวเลือก export ที่กำหนด
   *
   * @param request แถวข้อมูลและการตั้งค่ารูปแบบ export
   * @param output stream ปลายทาง ซึ่งจะยังเปิดอยู่หลังเขียนเสร็จ
   * @throws IOException เมื่อ POI เขียนเอกสาร XLSX ไม่สำเร็จ
   */
  public void writeTo(ExcelExportRequest<T> request, OutputStream output) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    writeRows(request.data(), request, output);
  }

  /** Exports a header-only workbook that callers can use as an upload template. */
  public ExcelExportResult exportTemplate(String sheetName, String filename) throws IOException {
    return export(new ExcelExportRequest<>(List.of(), sheetName, filename, null, null));
  }

  /**
   * Exports several worksheets with the same annotated row schema into one workbook. Each request
   * supplies its own sheet name and rows.
   */
  public ExcelExportResult exportWorkbook(List<ExcelExportRequest<T>> requests, String filename)
      throws IOException {
    try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      writeWorkbook(requests, output);
      return new ExcelExportResult(output.toByteArray(), filename);
    }
  }

  /** Writes a multi-sheet workbook directly to the supplied stream. */
  public void writeWorkbook(List<ExcelExportRequest<T>> requests, OutputStream output)
      throws IOException {
    Objects.requireNonNull(requests, "requests must not be null");
    Objects.requireNonNull(output, "output must not be null");
    if (requests.isEmpty()) throw new IllegalArgumentException("requests must not be empty");
    SXSSFWorkbook workbook = new SXSSFWorkbook(100);
    workbook.setCompressTempFiles(true);
    try (workbook) {
      for (ExcelExportRequest<T> request : requests) {
        writeSheet(
            workbook, Objects.requireNonNull(request, "request must not be null").data(), request);
      }
      workbook.write(output);
    }
  }

  /**
   * รวม workbook ที่เขียนแบบ streaming เป็น byte array เฉพาะเมื่อเรียก API ที่ต้องคืนค่า byte array
   */
  private ExcelExportResult exportRows(Iterable<T> data, ExcelExportRequest<T> request)
      throws IOException {
    try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      writeRows(data, request, output);
      return new ExcelExportResult(output.toByteArray(), request.filename());
    }
  }

  /** ตั้งค่า workbook แบบ streaming สร้าง sheet และเขียนผลลัพธ์ลง stream ของผู้เรียก */
  private void writeRows(Iterable<T> data, ExcelExportRequest<T> request, OutputStream output)
      throws IOException {
    Objects.requireNonNull(output, "output must not be null");
    SXSSFWorkbook workbook = new SXSSFWorkbook(100);
    workbook.setCompressTempFiles(true);
    try (workbook) {
      writeSheet(workbook, data, request);
      workbook.write(output);
    }
  }

  /** Adds one configured worksheet to a workbook already owned by the caller. */
  private void writeSheet(SXSSFWorkbook workbook, Iterable<T> data, ExcelExportRequest<T> request) {
    List<ExcelColumnDef> columns = columnsFor(getEntityClass());
    ExcelHeaderColumn[] headers =
        columns.stream()
            .map(
                column ->
                    ExcelHeaderColumn.create(column.name(), column.width(), column.headerColor()))
            .toArray(ExcelHeaderColumn[]::new);
    ExcelDataBuilderStyle style =
        new ExcelDataBuilderStyle(null, null, request.headerRowHeight(), request.fontFamily());
    var sheet =
        new ExcelDataBuilder(workbook, request.sheetName(), headers, mapRows(data), style, columns)
            .createSheet();
    if (request.freezeHeader()) sheet.createFreezePane(0, 1);
  }

  /** อ่านและตรวจสอบ schema ของคอลัมน์ตามลำดับที่ประกาศบน target type */
  private List<ExcelColumnDef> columnsFor(Class<?> type) {
    List<ExcelColumnDef> columns =
        java.util.Arrays.stream(type.getDeclaredFields())
            .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
            .map(ExcelColumnDef::fromField)
            .sorted(
                Comparator.comparingInt(
                    column -> column.index() < 0 ? Integer.MAX_VALUE : column.index()))
            .toList();
    if (columns.isEmpty()) {
      throw new IllegalArgumentException("No @ExcelColumn fields found on " + type.getName());
    }
    Set<Integer> indexes = new HashSet<>();
    for (ExcelColumnDef column : columns) {
      if (column.index() >= 0 && !indexes.add(column.index())) {
        throw new IllegalArgumentException(
            "Duplicate @ExcelColumn index " + column.index() + " on " + type.getName());
      }
    }
    return columns;
  }

  /** แปลงข้อมูลต้นทางเป็น array ของ cell แบบ lazy โดยไม่สร้างทุกแถวล่วงหน้า */
  private Iterable<Object[]> mapRows(Iterable<T> data) {
    return () ->
        new java.util.Iterator<>() {
          /** iterator ของข้อมูล domain ที่ผู้เรียกส่งมาเพื่อ export */
          private final java.util.Iterator<T> source = data.iterator();

          @Override
          public boolean hasNext() {
            return source.hasNext();
          }

          @Override
          public Object[] next() {
            return Objects.requireNonNull(mapToRow(source.next()), "mapToRow must not return null");
          }
        };
  }
}
