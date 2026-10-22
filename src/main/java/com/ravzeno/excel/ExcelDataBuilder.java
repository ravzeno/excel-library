package com.ravzeno.excel;

import com.ravzeno.excel.model.export.ExcelColumnDef;
import com.ravzeno.excel.model.export.ExcelHeaderColumn;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Builds an XLSX sheet from headers and row values. */
public final class ExcelDataBuilder {
  /** Workbook that receives the newly-created sheet; may be XSSF or streaming SXSSF. */
  private final Workbook workbook;

  /** Underlying XSSF workbook used because colors and styles are XSSF-specific objects. */
  private final XSSFWorkbook styleWorkbook;

  /** Name assigned to the sheet created by {@link #createSheet()}. */
  private final String sheetName;

  /** Ordered header definitions; their array positions define physical spreadsheet columns. */
  private final ExcelHeaderColumn[] headerColumns;

  /** Lazily-consumed source rows, allowing callers to stream a large export. */
  private final Iterable<Object[]> data;

  /** Reusable default style for header cells without column-specific options. */
  private final XSSFCellStyle headerStyle;

  /** Reusable default style for ordinary data cells without column-specific options. */
  private final XSSFCellStyle dataStyle;

  /** Height, in Excel twips, applied to the first/header row. */
  private final float headerRowHeight;

  /** Optional typeface applied to generated fonts. */
  private final String fontFamily;

  /** Maps physical zero-based sheet columns to their export configuration. */
  private final Map<Integer, ExcelColumnDef> columnConfigs;

  /** Deduplicates equivalent generated styles to stay below Excel's style limit. */
  private final Map<String, XSSFCellStyle> styles = new HashMap<>();

  /**
   * Creates a builder with default column styling.
   *
   * @param workbook target XLSX or streaming XLSX workbook
   * @param sheetName visible worksheet name
   * @param headerColumns ordered header definitions
   * @param data rows whose value order matches {@code headerColumns}
   * @param style shared style overrides
   */
  public ExcelDataBuilder(
      Workbook workbook,
      String sheetName,
      ExcelHeaderColumn[] headerColumns,
      Iterable<Object[]> data,
      ExcelDataBuilderStyle style) {
    this(workbook, sheetName, headerColumns, data, style, List.of());
  }

  /**
   * Creates a builder with per-column export configuration resolved from annotations.
   *
   * @param workbook target XLSX or streaming XLSX workbook
   * @param sheetName visible worksheet name
   * @param headerColumns ordered header definitions
   * @param data rows whose value order matches {@code headerColumns}
   * @param style shared style overrides
   * @param columnDefs ordered per-column formatting metadata; may be {@code null}
   */
  public ExcelDataBuilder(
      Workbook workbook,
      String sheetName,
      ExcelHeaderColumn[] headerColumns,
      Iterable<Object[]> data,
      ExcelDataBuilderStyle style,
      List<ExcelColumnDef> columnDefs) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
    this.styleWorkbook = getStyleWorkbook(workbook);
    this.sheetName = requireSheetName(sheetName);
    this.headerColumns =
        Objects.requireNonNull(headerColumns, "headerColumns must not be null").clone();
    this.data = Objects.requireNonNull(data, "data must not be null");
    this.fontFamily = style == null ? null : style.fontFamily();
    this.headerRowHeight =
        style == null || style.headerRowHeight() == null
            ? ExcelDataBuilderStyle.DEFAULT_HEADER_ROW_HEIGHT
            : style.headerRowHeight();
    this.headerStyle =
        style == null || style.headerStyle() == null ? defaultHeaderStyle() : style.headerStyle();
    this.dataStyle =
        style == null || style.dataStyle() == null ? defaultDataStyle() : style.dataStyle();
    this.columnConfigs = new HashMap<>();
    if (columnDefs != null) {
      for (int position = 0; position < columnDefs.size(); position++) {
        this.columnConfigs.put(position, columnDefs.get(position));
      }
    }
  }

  /**
   * Creates the header and writes each source row to a new worksheet.
   *
   * @return the created worksheet, owned by the supplied workbook
   * @throws IllegalArgumentException if a row has more values than configured columns
   */
  public Sheet createSheet() {
    Sheet sheet = workbook.createSheet(sheetName);
    createHeaderRow(sheet);
    int rowIndex = 1;
    for (Object[] values : data) {
      Object[] nonNullValues = Objects.requireNonNull(values, "data rows must not be null");
      if (nonNullValues.length > headerColumns.length) {
        throw new IllegalArgumentException(
            "A data row has "
                + nonNullValues.length
                + " values, but only "
                + headerColumns.length
                + " columns are defined");
      }
      Row row = sheet.createRow(rowIndex++);
      for (int column = 0; column < nonNullValues.length; column++)
        addDataCell(row, column, nonNullValues[column]);
    }
    addDropdowns(sheet, rowIndex);
    return sheet;
  }

  /**
   * Adds annotation-declared dropdown validation below the header, including future template rows.
   */
  private void addDropdowns(Sheet sheet, int nextRow) {
    for (Map.Entry<Integer, ExcelColumnDef> entry : columnConfigs.entrySet()) {
      String[] values = entry.getValue().allowedValues();
      if (values == null || values.length == 0) continue;
      DataValidationConstraint constraint =
          sheet.getDataValidationHelper().createExplicitListConstraint(values);
      int lastRow = Math.max(nextRow + 999, 1_000);
      DataValidation validation =
          sheet
              .getDataValidationHelper()
              .createValidation(
                  constraint, new CellRangeAddressList(1, lastRow, entry.getKey(), entry.getKey()));
      validation.setSuppressDropDownArrow(false);
      validation.setShowErrorBox(true);
      sheet.addValidationData(validation);
    }
  }

  /** Creates row zero, including header text, widths, and column-specific header styling. */
  private void createHeaderRow(Sheet sheet) {
    Row row = sheet.createRow(0);
    row.setHeight((short) headerRowHeight);
    for (int index = 0; index < headerColumns.length; index++) {
      ExcelHeaderColumn column = headerColumns[index];
      Cell cell = row.createCell(index);
      cell.setCellValue(column.name());
      cell.setCellStyle(headerStyle(column, columnConfigs.get(index)));
      sheet.setColumnWidth(index, column.width());
    }
  }

  /** Converts a Java value to its Excel cell representation and applies the matching style. */
  private void addDataCell(Row row, int index, Object value) {
    Cell cell = row.createCell(index);
    ExcelColumnDef config = columnConfigs.get(index);
    cell.setCellStyle(dataStyle(config, value));
    if (value == null) cell.setBlank();
    else if (value instanceof BigDecimal decimal) cell.setCellValue(decimal.doubleValue());
    else if (value instanceof Number number) cell.setCellValue(number.doubleValue());
    else if (value instanceof Boolean bool) cell.setCellValue(bool);
    else if (value instanceof LocalDate date) cell.setCellValue(date);
    else if (value instanceof LocalDateTime dateTime) cell.setCellValue(dateTime);
    else if (value instanceof Instant instant)
      cell.setCellValue(LocalDateTime.ofInstant(instant, ZoneId.systemDefault()));
    else cell.setCellValue(value.toString());
  }

  /** Resolves a cached header style, creating a colored override only when needed. */
  private XSSFCellStyle headerStyle(ExcelHeaderColumn header, ExcelColumnDef config) {
    if (header.headerColor() == null && config == null) return headerStyle;
    XSSFColor background = header.headerColor() == null ? rgb(0, 47, 83) : header.headerColor();
    String key =
        "header:"
            + background.getARGBHex()
            + ":"
            + (config == null
                ? ""
                : config.bold() + ":" + color(config.fontColor()) + ":" + config.fontSize());
    return styles.computeIfAbsent(key, ignored -> createStyle(background, config, true, null));
  }

  /** Resolves a cached data style using the value type to select date or number formatting. */
  private XSSFCellStyle dataStyle(ExcelColumnDef config, Object value) {
    if (config == null) return dataStyle;
    boolean date =
        value instanceof LocalDate || value instanceof LocalDateTime || value instanceof Instant;
    boolean numeric = value instanceof Number;
    String key =
        "data:"
            + color(config.dataColor())
            + ":"
            + config.dataBold()
            + ":"
            + color(config.dataFontColor())
            + ":"
            + config.dataFontSize()
            + ":"
            + config.align()
            + ":"
            + (date ? config.dateFormat() : "")
            + ":"
            + (numeric ? config.numberFormat() : "");
    return styles.computeIfAbsent(
        key,
        ignored ->
            createStyle(
                config.dataColor(),
                config,
                false,
                date
                    ? config.dateFormat().getPattern()
                    : numeric ? config.numberFormat().getPattern() : null));
  }

  /** Builds the dark-blue, white-text default header style. */
  private XSSFCellStyle defaultHeaderStyle() {
    XSSFCellStyle style = baseStyle();
    XSSFFont font = styleWorkbook.createFont();
    applyFont(font);
    font.setBold(true);
    font.setColor(IndexedColors.WHITE.getIndex());
    style.setFont(font);
    style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  /** Builds the centered, bordered default data style. */
  private XSSFCellStyle defaultDataStyle() {
    XSSFCellStyle style = baseStyle();
    if (fontFamily != null) {
      XSSFFont font = styleWorkbook.createFont();
      applyFont(font);
      style.setFont(font);
    }
    return style;
  }

  /** Builds a style from a fill color, font settings, alignment, and optional Excel format. */
  private XSSFCellStyle createStyle(
      XSSFColor fill, ExcelColumnDef config, boolean header, String format) {
    XSSFCellStyle style = baseStyle();
    if (fill != null) {
      style.setFillForegroundColor(fill);
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
    if (config != null && config.align() != null && !header) style.setAlignment(config.align());
    XSSFFont font = styleWorkbook.createFont();
    applyFont(font);
    if (header) {
      font.setBold(config == null || config.bold());
      if (config == null || config.fontColor() == null) {
        font.setColor(IndexedColors.WHITE.getIndex());
      } else {
        font.setColor(config.fontColor());
      }
      if (config != null && config.fontSize() != null)
        font.setFontHeightInPoints(config.fontSize().shortValue());
    } else {
      font.setBold(config.dataBold());
      if (config.dataFontColor() != null) font.setColor(config.dataFontColor());
      if (config.dataFontSize() != null)
        font.setFontHeightInPoints(config.dataFontSize().shortValue());
    }
    style.setFont(font);
    if (format != null) style.setDataFormat(styleWorkbook.createDataFormat().getFormat(format));
    return style;
  }

  /** Creates the common bordered and vertically-centered foundation for all generated styles. */
  private XSSFCellStyle baseStyle() {
    XSSFCellStyle style = styleWorkbook.createCellStyle();
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setAlignment(HorizontalAlignment.CENTER);
    style.setVerticalAlignment(VerticalAlignment.CENTER);
    return style;
  }

  /** Applies the optional configured typeface to a newly-created font. */
  private void applyFont(XSSFFont font) {
    if (fontFamily != null) font.setFontName(fontFamily);
  }

  /** Produces a stable color key for the style cache. */
  private static String color(XSSFColor color) {
    return color == null ? "" : color.getARGBHex();
  }

  /** Creates an XSSF RGB color from standard 0-255 components. */
  private static XSSFColor rgb(int red, int green, int blue) {
    return new XSSFColor(new byte[] {(byte) red, (byte) green, (byte) blue}, null);
  }

  /** Rejects worksheet names that Apache POI cannot create meaningfully. */
  private static String requireSheetName(String sheetName) {
    if (sheetName == null || sheetName.isBlank())
      throw new IllegalArgumentException("sheetName must not be blank");
    return sheetName;
  }

  /** Finds the backing XSSF workbook used to create XSSF-compatible styles for either mode. */
  private static XSSFWorkbook getStyleWorkbook(Workbook workbook) {
    if (workbook instanceof XSSFWorkbook xssfWorkbook) {
      return xssfWorkbook;
    }
    if (workbook instanceof SXSSFWorkbook streamingWorkbook) {
      return streamingWorkbook.getXSSFWorkbook();
    }
    throw new IllegalArgumentException("Only XSSFWorkbook and SXSSFWorkbook are supported");
  }
}
    /**
     * Creates a builder with annotation-derived, per-column export options.
     *
     * @param columnDefs ordered column definitions; positions, rather than annotation indexes,
     *     correspond to physical spreadsheet columns
     */
