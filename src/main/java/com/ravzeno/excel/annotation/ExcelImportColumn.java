package com.ravzeno.excel.annotation;

import com.ravzeno.excel.enums.ExcelDataType;
import com.ravzeno.excel.enums.ExcelDateFormat;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Declares how one DTO field receives and validates a column during Excel import. The {@link
 * #index()} value is zero-based because it maps to the physical Excel column.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImportColumn {
  /**
   * Returns the optional expected header text; non-blank values are validated against the uploaded
   * header row.
   *
   * @return expected header text
   */
  String headerName() default "";

  /**
   * Returns the zero-based Excel column index from which this field is read.
   *
   * @return physical column index
   */
  int index();

  /**
   * Returns the expected value type; {@code AUTO} selects conversion from the target field type.
   *
   * @return expected value type
   */
  ExcelDataType dataType() default ExcelDataType.AUTO;

  /**
   * Returns the date pattern tried first when reading a date value.
   *
   * @return preferred date format
   */
  ExcelDateFormat dateFormat() default ExcelDateFormat.THAI_DATE;

  /**
   * Returns whether an empty cell is reported as a validation error.
   *
   * @return required flag
   */
  boolean required() default false;

  /**
   * Returns the minimum text length; {@code -1} disables the constraint.
   *
   * @return minimum text length
   */
  int minLength() default -1;

  /**
   * Returns the maximum text length; {@code -1} disables the constraint.
   *
   * @return maximum text length
   */
  int maxLength() default -1;

  /**
   * Returns the minimum allowed numeric value; {@link Double#NaN} disables the constraint.
   *
   * @return inclusive numeric minimum
   */
  double minValue() default Double.NaN;

  /**
   * Returns the maximum allowed numeric value; {@link Double#NaN} disables the constraint.
   *
   * @return inclusive numeric maximum
   */
  double maxValue() default Double.NaN;

  /**
   * Returns permitted text values; an empty array disables the constraint.
   *
   * @return permitted values
   */
  String[] allowedValues() default {};

  /**
   * Returns the Java regular expression the cell text must match; blank disables the constraint.
   *
   * @return validation regular expression
   */
  String pattern() default "";
}
