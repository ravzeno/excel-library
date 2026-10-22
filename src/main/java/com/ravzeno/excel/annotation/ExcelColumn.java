package com.ravzeno.excel.annotation;

import com.ravzeno.excel.enums.ExcelDateFormat;
import com.ravzeno.excel.enums.ExcelNumberFormat;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Declares how one DTO field is rendered as a column during Excel export. Fields without this
 * annotation are ignored by the export mapper.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
  /**
   * Returns the header text shown in the generated spreadsheet.
   *
   * @return header text
   */
  String name();

  /**
   * Returns the explicit source column index; {@code -1} appends this field after indexed columns.
   *
   * @return source column index
   */
  int index() default -1;

  /**
   * Returns the column width in Excel character-width units; {@code -1} uses the default width.
   *
   * @return requested column width
   */
  int width() default -1;

  /**
   * Returns the header background colour as {@code #RRGGBB}; {@code NONE} means no fill.
   *
   * @return header fill colour
   */
  String headerCellColor() default "NONE";

  /**
   * Returns whether the header font is bold.
   *
   * @return header bold flag
   */
  boolean headerFontBold() default false;

  /**
   * Returns the header font colour as {@code #RRGGBB}; {@code NONE} keeps the default colour.
   *
   * @return header font colour
   */
  String headerFontColor() default "NONE";

  /**
   * Returns the header font size in points; a non-positive value leaves the size unspecified.
   *
   * @return header font size
   */
  int headerFontSize() default 16;

  /**
   * Returns the data-cell background colour as {@code #RRGGBB}; {@code NONE} means no fill.
   *
   * @return data-cell fill colour
   */
  String dataColor() default "NONE";

  /**
   * Returns whether the data-cell font is bold.
   *
   * @return data-cell bold flag
   */
  boolean dataBold() default false;

  /**
   * Returns the data-cell font colour as {@code #RRGGBB}; {@code NONE} keeps the default colour.
   *
   * @return data-cell font colour
   */
  String dataFontColor() default "NONE";

  /**
   * Returns the data-cell font size in points; a non-positive value leaves the size unspecified.
   *
   * @return data-cell font size
   */
  int dataFontSize() default 16;

  /**
   * Returns the horizontal alignment used for data cells.
   *
   * @return data-cell alignment
   */
  HorizontalAlignment align() default HorizontalAlignment.CENTER;

  /**
   * Returns the date pattern used when the field value is a date or date-time type.
   *
   * @return date format
   */
  ExcelDateFormat dateFormat() default ExcelDateFormat.THAI_DATE;

  /**
   * Returns the number pattern used when the field value is numeric.
   *
   * @return number format
   */
  ExcelNumberFormat numberFormat() default ExcelNumberFormat.DECIMAL_2;

  /**
   * Returns permitted values shown as a dropdown in generated templates; empty disables it.
   *
   * @return allowed cell values
   */
  String[] allowedValues() default {};
}
