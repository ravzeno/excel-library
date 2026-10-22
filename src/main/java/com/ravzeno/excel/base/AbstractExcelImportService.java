package com.ravzeno.excel.base;

import com.ravzeno.excel.annotation.ExcelImportColumn;
import com.ravzeno.excel.enums.ExcelDataType;
import com.ravzeno.excel.enums.ExcelImportProblem;
import com.ravzeno.excel.enums.ExcelImportProblemCode;
import com.ravzeno.excel.model.ConvertValueData;
import com.ravzeno.excel.model.ExcelReadDataDto;
import com.ravzeno.excel.model.import_.ExcelImportError;
import com.ravzeno.excel.model.import_.ExcelImportRequest;
import com.ravzeno.excel.model.import_.ExcelImportResult;
import com.ravzeno.excel.util.ExcelUtils;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

/**
 * คลาสพื้นฐานสำหรับ import Excel โดยกำหนดการ map และ validation จาก annotation
 *
 * @param <T> type ปลายทาง ซึ่งต้องมี constructor เปล่าที่เข้าถึงได้
 */
public abstract class AbstractExcelImportService<T> {
  /** constructor สำหรับคลาสลูก */
  protected AbstractExcelImportService() {}

  /**
   * Creates an annotation-based importer without requiring a dedicated subclass.
   *
   * @param entityClass type that declares {@code @ExcelImportColumn} fields
   * @param sheetNo one-based worksheet number
   * @param headerRow one-based header row number
   * @param dataRow one-based first data row number
   * @param <T> target row type
   * @return a ready-to-use importer
   */
  public static <T> AbstractExcelImportService<T> of(
      Class<T> entityClass, int sheetNo, int headerRow, int dataRow) {
    Objects.requireNonNull(entityClass, "entityClass must not be null");
    return new AbstractExcelImportService<>() {
      @Override
      protected int getSheetNo() {
        return sheetNo;
      }

      @Override
      protected int getHeaderRow() {
        return headerRow;
      }

      @Override
      protected int getDataRow() {
        return dataRow;
      }

      @Override
      protected Class<T> getEntityClass() {
        return entityClass;
      }
    };
  }

  /** ค่าพิเศษสำหรับ error ที่เกิดกับทั้งแถว ไม่ได้ผูกกับคอลัมน์ใดคอลัมน์หนึ่ง */
  private static final int NO_COLUMN = -1;

  /** cache metadata จาก reflection เพราะทุกแถวใช้ annotation ชุดเดียวกัน */
  private static final Map<Class<?>, List<FieldColumnMapping>> FIELD_MAPPING_CACHE =
      new ConcurrentHashMap<>();

  /**
   * ระบุหมายเลข worksheet แบบเริ่มที่ 1 เมื่อผู้เรียกไม่ได้ส่ง request
   *
   * @return หมายเลข worksheet เริ่มต้น
   */
  protected abstract int getSheetNo();

  /**
   * ระบุหมายเลขแถว header แบบเริ่มที่ 1 เมื่อผู้เรียกไม่ได้ส่ง request
   *
   * @return หมายเลขแถว header เริ่มต้น
   */
  protected abstract int getHeaderRow();

  /**
   * ระบุหมายเลขแถวข้อมูลแรกแบบเริ่มที่ 1 เมื่อผู้เรียกไม่ได้ส่ง request
   *
   * @return หมายเลขแถวข้อมูลแรกเริ่มต้น
   */
  protected abstract int getDataRow();

  /**
   * ระบุ type ที่ติด annotation ซึ่งจะสร้างขึ้นสำหรับทุกแถวที่ valid
   *
   * @return DTO type ที่ติด annotation
   */
  protected abstract Class<T> getEntityClass();

  /**
   * ตรวจสอบ header ก่อน map แถว สามารถ override เพื่อบังคับ header ให้ตรงตาม template
   *
   * @param headers ค่า header ที่อ่านจากแถว worksheet ที่กำหนด
   */
  protected void validateHeaders(List<String> headers) {}

  /**
   * Performs optional domain-specific validation after a row has been converted successfully.
   * Return errors using zero-based worksheet coordinates; a non-empty result rejects that row.
   *
   * @param value mapped row
   * @param row zero-based physical worksheet row
   * @return additional row or cell errors
   */
  protected List<ExcelImportError> validateRow(T value, int row) {
    return List.of();
  }

  /**
   * Import เอกสาร XLSX จาก byte โดยใช้ layout เริ่มต้นของคลาสลูก
   *
   * @param bytes เนื้อหา binary ของ XLSX ฉบับเต็ม
   * @return แถวที่ map สำเร็จพร้อม validation error
   * @throws IOException เมื่ออ่านเอกสาร XLSX ไม่สำเร็จ
   */
  public ExcelImportResult<T> importFromExcel(byte[] bytes) throws IOException {
    return importFromExcel(bytes, null);
  }

  /**
   * Import เอกสาร XLSX จาก byte โดยใช้ layout worksheet ที่ระบุ
   *
   * @param bytes เนื้อหา binary ของ XLSX ฉบับเต็ม
   * @param request การตั้งค่า sheet, header และแถวข้อมูลแรก; {@code null} ใช้ค่าเริ่มต้นของคลาสลูก
   * @return แถวที่ map สำเร็จพร้อม validation error
   * @throws IOException เมื่ออ่านเอกสาร XLSX ไม่สำเร็จ
   */
  public ExcelImportResult<T> importFromExcel(byte[] bytes, ExcelImportRequest request)
      throws IOException {
    try (InputStream input =
        new ByteArrayInputStream(Objects.requireNonNull(bytes, "bytes must not be null"))) {
      return importFromExcel(input, request);
    }
  }

  /**
   * Imports an XLSX document from a stream. The stream is consumed and closed before this method
   * returns.
   *
   * @param input XLSX source stream
   * @return valid mapped rows and validation errors
   * @throws IOException if the XLSX document cannot be parsed
   */
  public ExcelImportResult<T> importFromExcel(InputStream input) throws IOException {
    return importFromExcel(input, null);
  }

  /**
   * Imports an XLSX document from a stream with the supplied worksheet layout.
   *
   * @param input XLSX source stream; consumed but not closed by this method
   * @param request sheet, header, and first data row settings; {@code null} uses subclass defaults
   * @return valid mapped rows and validation errors
   * @throws IOException if the XLSX document cannot be parsed
   */
  public ExcelImportResult<T> importFromExcel(InputStream input, ExcelImportRequest request)
      throws IOException {
    Objects.requireNonNull(input, "input must not be null");
    ExcelImportRequest resolved = resolveRequest(request);
    return readExcel(input, resolved);
  }

  /** Combines an optional per-call layout with the subclass's default worksheet layout. */
  private ExcelImportRequest resolveRequest(ExcelImportRequest request) {
    return request == null
        ? new ExcelImportRequest(getSheetNo(), getHeaderRow(), getDataRow())
        : request;
  }

  /**
   * Parses the requested worksheet through POI's SAX API, keeping only the current XML row in
   * memory while mapping it to a target object.
   */
  @SuppressWarnings(
      "deprecation") // POI's streaming shared-string reader avoids loading strings into memory.
  private ExcelImportResult<T> readExcel(InputStream input, ExcelImportRequest request)
      throws IOException {
    int sheetNo = request.getSheetNoOrDefault();
    if (sheetNo < 1) {
      throw new IllegalArgumentException("Sheet number must be at least 1, got: " + sheetNo);
    }
    int headerRow = request.getHeaderRowOrDefault();
    int dataRow = request.getDataRowOrDefault();
    if (headerRow < 1) {
      throw new IllegalArgumentException("Header row must be at least 1, got: " + headerRow);
    }
    if (dataRow <= headerRow) {
      throw new IllegalArgumentException(
          "Data row must be after header row, got header row "
              + headerRow
              + " and data row "
              + dataRow);
    }
    StreamingSheetHandler handler =
        new StreamingSheetHandler(request, mappingsFor(getEntityClass()));
    try (OPCPackage packageFile = OPCPackage.open(input)) {
      XSSFReader reader = new XSSFReader(packageFile);
      try (InputStream sheet = sheetFor(reader, sheetNo)) {
        XMLReader parser = SAXHelper.newXMLReader();
        parser.setContentHandler(
            new XSSFSheetXMLHandler(
                reader.getStylesTable(),
                null,
                new ReadOnlySharedStringsTable(packageFile),
                handler,
                new DataFormatter(),
                false));
        parser.parse(new InputSource(sheet));
      }
    } catch (IllegalArgumentException exception) {
      throw exception;
    } catch (Exception exception) {
      throw new IOException("Unable to parse XLSX worksheet", exception);
    }
    return handler.toResult();
  }

  /** Selects the requested one-based sheet and closes all preceding sheet streams. */
  private InputStream sheetFor(XSSFReader reader, int sheetNo)
      throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
    Iterator<InputStream> sheets = reader.getSheetsData();
    for (int currentSheet = 1; sheets.hasNext(); currentSheet++) {
      InputStream sheet = sheets.next();
      if (currentSheet == sheetNo) {
        return sheet;
      }
      sheet.close();
    }
    throw new IllegalArgumentException("Sheet number does not exist: " + sheetNo);
  }

  /** Converts one logical worksheet row into a target instance, recording validation errors. */
  private T mapRow(
      List<ExcelReadDataDto> row,
      List<FieldColumnMapping> mappings,
      int rowNumber,
      List<ExcelImportError> errors) {
    try {
      T target = instantiate();
      for (FieldColumnMapping mapping : mappings) {
        int index = mapping.annotation().index();
        String raw = index < row.size() ? row.get(index).value() : null;
        ConvertValueData converted = convert(raw, mapping);
        if (!validate(converted, raw, mapping, rowNumber, errors)) return null;
        if (converted.value() != null || !mapping.field().getType().isPrimitive()) {
          mapping.field().set(target, converted.value());
        }
      }
      return target;
    } catch (ReflectiveOperationException | IllegalArgumentException e) {
      addError(
          errors,
          rowNumber,
          NO_COLUMN,
          null,
          null,
          ExcelImportProblemCode.MAPPING_ERROR,
          e.getMessage());
      return null;
    }
  }

  private final class StreamingSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    /** Layout used to identify header and first data rows. */
    private final ExcelImportRequest request;

    /** Cached annotated fields used to map the current row. */
    private final List<FieldColumnMapping> mappings;

    /** Successfully mapped values returned to the caller after parsing completes. */
    private final List<T> values = new ArrayList<>();

    /** Per-row validation and conversion errors returned to the caller. */
    private final List<ExcelImportError> errors = new ArrayList<>();

    /** Header labels read from the configured header row. */
    private List<String> headers = List.of();

    /** Sparse cell values for the XML row currently being parsed. */
    private List<String> currentRow = List.of();

    /** Whether the expected header row was encountered in the input sheet. */
    private boolean readHeaders;

    /** Whether annotation-declared headers agree with the uploaded worksheet. */
    private boolean headersValid = true;

    /** Count of non-empty data rows processed, regardless of validation result. */
    private int totalRows;

    private StreamingSheetHandler(ExcelImportRequest request, List<FieldColumnMapping> mappings) {
      this.request = request;
      this.mappings = mappings;
    }

    @Override
    public void startRow(int rowNumber) {
      currentRow = new ArrayList<>();
    }

    @Override
    public void endRow(int rowNumber) {
      if (rowNumber == request.getHeaderRowOrDefault() - 1) {
        headers = List.copyOf(currentRow);
        readHeaders = true;
        headersValid = validateAnnotatedHeaders();
        validateHeaders(headers);
        return;
      }
      if (!readHeaders
          || !headersValid
          || rowNumber < request.getDataRowOrDefault() - 1
          || isCurrentRowEmpty()) {
        return;
      }

      totalRows++;
      T mapped = mapRow(toReadDataRow(), mappings, rowNumber, errors);
      if (mapped != null) {
        List<ExcelImportError> customErrors = validateRow(mapped, rowNumber);
        if (customErrors == null) {
          throw new IllegalStateException("validateRow must not return null");
        }
        if (customErrors.isEmpty()) values.add(mapped);
        else errors.addAll(customErrors);
      }
    }

    @Override
    public void cell(
        String cellReference,
        String formattedValue,
        org.apache.poi.xssf.usermodel.XSSFComment comment) {
      int columnIndex = new CellReference(cellReference).getCol();
      while (currentRow.size() <= columnIndex) {
        currentRow.add("");
      }
      currentRow.set(columnIndex, formattedValue == null ? "" : formattedValue);
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {}

    private ExcelImportResult<T> toResult() {
      if (!readHeaders) {
        errors.add(
            new ExcelImportError(
                request.getHeaderRowOrDefault() - 1,
                NO_COLUMN,
                null,
                null,
                ExcelImportProblemCode.HEADER_ROW_NOT_FOUND,
                ExcelImportProblemCode.HEADER_ROW_NOT_FOUND.getMessage()));
        validateHeaders(List.of());
      }
      int firstDataRow = request.getDataRowOrDefault() - 1;
      int invalidRows =
          (int)
              errors.stream()
                  .map(ExcelImportError::row)
                  .filter(row -> row >= firstDataRow)
                  .distinct()
                  .count();
      return new ExcelImportResult<>(values, errors, totalRows, values.size(), invalidRows);
    }

    private boolean isCurrentRowEmpty() {
      return currentRow.stream().allMatch(String::isBlank);
    }

    private List<ExcelReadDataDto> toReadDataRow() {
      List<ExcelReadDataDto> row = new ArrayList<>(headers.size());
      for (int column = 0; column < headers.size(); column++) {
        String value = column < currentRow.size() ? currentRow.get(column) : "";
        row.add(new ExcelReadDataDto(value));
      }
      return row;
    }

    /** Checks every non-blank annotation header before accepting any data rows. */
    private boolean validateAnnotatedHeaders() {
      boolean valid = true;
      for (FieldColumnMapping mapping : mappings) {
        ExcelImportColumn annotation = mapping.annotation();
        if (annotation.headerName().isBlank()) {
          continue;
        }
        int index = annotation.index();
        String actual = index < headers.size() ? headers.get(index) : "";
        if (!annotation.headerName().trim().equals(actual.trim())) {
          addError(
              errors,
              request.getHeaderRowOrDefault() - 1,
              index,
              mapping,
              actual,
              ExcelImportProblemCode.HEADER_MISMATCH,
              annotation.headerName(),
              actual);
          valid = false;
        }
      }
      return valid;
    }
  }

  /** Instantiates the user DTO through its accessible no-arguments constructor. */
  private T instantiate() throws ReflectiveOperationException {
    Constructor<T> constructor = getEntityClass().getDeclaredConstructor();
    constructor.setAccessible(true);
    return constructor.newInstance();
  }

  /** Converts one formatted Excel cell string according to its annotation and field type. */
  private ConvertValueData convert(String raw, FieldColumnMapping mapping) {
    if (raw == null || raw.isBlank()) {
      return new ConvertValueData(null, true);
    }
    try {
      ExcelDataType type = mapping.annotation().dataType();
      Class<?> target = mapping.field().getType();
      return switch (type) {
        case STRING ->
            supportsString(target)
                ? new ConvertValueData(raw.trim(), true)
                : ConvertValueData.failure();
        case DATE -> convertDate(raw, target, mapping);
        case NUMERIC -> convertNumber(raw, target);
        case AUTO -> convertAutomatically(raw, target, mapping);
      };
    } catch (RuntimeException exception) {
      return ConvertValueData.failure();
    }
  }

  /** Infers an appropriate value type for fields declared with {@link ExcelDataType#AUTO}. */
  private ConvertValueData convertAutomatically(
      String raw, Class<?> target, FieldColumnMapping mapping) {
    if (target == String.class) {
      return new ConvertValueData(raw.trim(), true);
    }
    if (target == LocalDate.class) {
      return convertDate(raw, target, mapping);
    }
    if (target == LocalDateTime.class || target == Instant.class) {
      return convertDate(raw, target, mapping);
    }
    if (target.isEnum()) {
      return convertEnum(raw, target);
    }
    if (target == Object.class) {
      if (ExcelUtils.isNumeric(raw)) {
        return convertNumber(raw, target);
      }
      if (ExcelUtils.isDate(raw)) {
        return convertDate(raw, target, mapping);
      }
      return new ConvertValueData(raw.trim(), true);
    }
    return convertNumber(raw, target);
  }

  /** Parses a date cell, treating conventional spreadsheet placeholders as an optional null. */
  private ConvertValueData convertDate(String raw, Class<?> target, FieldColumnMapping mapping) {
    if (!supportsDate(target)) {
      return ConvertValueData.failure();
    }
    if ("-".equals(raw) || "N/A".equalsIgnoreCase(raw)) {
      return new ConvertValueData(null, true);
    }
    if (target == LocalDate.class || target == Object.class) {
      LocalDate value = ExcelUtils.parseDate(raw, mapping.annotation().dateFormat());
      return value == null ? ConvertValueData.failure() : new ConvertValueData(value, true);
    }
    LocalDateTime value = ExcelUtils.parseDateTime(raw, mapping.annotation().dateFormat());
    if (value == null) return ConvertValueData.failure();
    return new ConvertValueData(
        target == Instant.class ? value.atZone(ZoneId.systemDefault()).toInstant() : value, true);
  }

  /** Matches enum names case-insensitively so spreadsheet values remain user-friendly. */
  @SuppressWarnings({"unchecked", "rawtypes"})
  private ConvertValueData convertEnum(String raw, Class<?> target) {
    for (Object constant : target.getEnumConstants()) {
      if (((Enum) constant).name().equalsIgnoreCase(raw.trim())) {
        return new ConvertValueData(constant, true);
      }
    }
    return ConvertValueData.failure();
  }

  /** Parses a decimal and narrows it only when the declared field type supports that conversion. */
  private ConvertValueData convertNumber(String raw, Class<?> target) {
    if (!supportsNumber(target) || !ExcelUtils.isNumeric(raw)) {
      return ConvertValueData.failure();
    }
    BigDecimal value = new BigDecimal(raw.replace(",", "").trim());
    return new ConvertValueData(numberForTarget(value, target), true);
  }

  /** Converts an exact decimal to the declared Java number type. */
  private Object numberForTarget(BigDecimal value, Class<?> type) {
    if (type == BigDecimal.class || type == Object.class) return value;
    if (type == Integer.class || type == int.class) return value.intValueExact();
    if (type == Long.class || type == long.class) return value.longValueExact();
    if (type == Double.class || type == double.class) return value.doubleValue();
    if (type == Float.class || type == float.class) return value.floatValue();
    return value.shortValueExact();
  }

  private boolean supportsString(Class<?> type) {
    return type == String.class || type == Object.class;
  }

  private boolean supportsDate(Class<?> type) {
    return type == LocalDate.class
        || type == LocalDateTime.class
        || type == Instant.class
        || type == Object.class;
  }

  private boolean supportsNumber(Class<?> type) {
    return type == BigDecimal.class
        || type == Object.class
        || type == Integer.class
        || type == int.class
        || type == Long.class
        || type == long.class
        || type == Double.class
        || type == double.class
        || type == Float.class
        || type == float.class
        || type == Short.class
        || type == short.class;
  }

  /** Applies required, length, and regular-expression constraints after conversion. */
  private boolean validate(
      ConvertValueData value,
      String raw,
      FieldColumnMapping mapping,
      int row,
      List<ExcelImportError> errors) {
    ExcelImportColumn annotation = mapping.annotation();
    String text = value.value() == null ? (raw == null ? "" : raw) : value.value().toString();
    if (annotation.required()
        && (!value.isConvertSuccess() || value.value() == null || text.isBlank())) {
      addError(
          errors, row, annotation.index(), mapping, raw, ExcelImportProblemCode.REQUIRED_FIELD);
      return false;
    }
    if (!value.isConvertSuccess()) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.INVALID_FORMAT,
          annotation.dataType().name().toLowerCase());
      return false;
    }
    if (value.value() == null) {
      return true;
    }
    if (annotation.minLength() > 0 && text.length() < annotation.minLength()) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.MIN_LENGTH,
          annotation.minLength());
      return false;
    }
    if (annotation.maxLength() > 0 && text.length() > annotation.maxLength()) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.MAX_LENGTH,
          annotation.maxLength());
      return false;
    }
    if (!Double.isNaN(annotation.minValue())
        && value.value() instanceof Number number
        && Double.compare(number.doubleValue(), annotation.minValue()) < 0) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.MIN_VALUE,
          annotation.minValue());
      return false;
    }
    if (!Double.isNaN(annotation.maxValue())
        && value.value() instanceof Number number
        && Double.compare(number.doubleValue(), annotation.maxValue()) > 0) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.MAX_VALUE,
          annotation.maxValue());
      return false;
    }
    if (annotation.allowedValues().length > 0
        && java.util.Arrays.stream(annotation.allowedValues()).noneMatch(text::equals)) {
      addError(
          errors, row, annotation.index(), mapping, raw, ExcelImportProblemCode.VALUE_NOT_ALLOWED);
      return false;
    }
    return matchesPattern(annotation, text, row, raw, mapping, errors);
  }

  /** Validates an optional regular-expression constraint and records its precise failure. */
  private boolean matchesPattern(
      ExcelImportColumn annotation,
      String text,
      int row,
      String raw,
      FieldColumnMapping mapping,
      List<ExcelImportError> errors) {
    if (annotation.pattern().isEmpty()) {
      return true;
    }
    try {
      if (Pattern.compile(annotation.pattern()).matcher(text).matches()) {
        return true;
      }
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.PATTERN_MISMATCH,
          annotation.pattern());
    } catch (PatternSyntaxException exception) {
      addError(
          errors,
          row,
          annotation.index(),
          mapping,
          raw,
          ExcelImportProblemCode.INVALID_PATTERN_REGEX,
          annotation.pattern());
    }
    return false;
  }

  /** Builds the field-to-column schema once per DTO class and caches it for later imports. */
  private static List<FieldColumnMapping> mappingsFor(Class<?> type) {
    return FIELD_MAPPING_CACHE.computeIfAbsent(
        type,
        key -> {
          List<FieldColumnMapping> mappings =
              java.util.Arrays.stream(key.getDeclaredFields())
                  .filter(field -> field.isAnnotationPresent(ExcelImportColumn.class))
                  .peek(field -> field.setAccessible(true))
                  .map(
                      field ->
                          new FieldColumnMapping(
                              field, field.getAnnotation(ExcelImportColumn.class)))
                  .sorted(Comparator.comparingInt(mapping -> mapping.annotation().index()))
                  .toList();
          if (mappings.isEmpty()) {
            throw new IllegalArgumentException(
                "No @ExcelImportColumn fields found on " + key.getName());
          }
          boolean duplicateIndexes =
              mappings.stream().map(mapping -> mapping.annotation().index()).distinct().count()
                  != mappings.size();
          if (duplicateIndexes) {
            throw new IllegalArgumentException(
                "Duplicate @ExcelImportColumn index on " + key.getName());
          }
          if (mappings.getFirst().annotation().index() < 0) {
            throw new IllegalArgumentException(
                "@ExcelImportColumn index must not be negative on " + key.getName());
          }
          return mappings;
        });
  }

  /** Creates a structured import error with the spreadsheet location and human-readable message. */
  private void addError(
      List<ExcelImportError> errors,
      int row,
      int column,
      FieldColumnMapping mapping,
      String value,
      ExcelImportProblem problem,
      Object... arguments) {
    String name =
        mapping == null
            ? null
            : mapping.annotation().headerName().isBlank()
                ? mapping.field().getName()
                : mapping.annotation().headerName();
    String message =
        arguments.length == 0
            ? problem.getMessage()
            : String.format(problem.getMessage(), arguments);
    errors.add(new ExcelImportError(row, column, name, value, problem, message));
  }

  private record FieldColumnMapping(Field field, ExcelImportColumn annotation) {}
}
