package com.ravzeno.excel.enums;

/** Built-in validation and mapping errors returned by the import pipeline. */
public enum ExcelImportProblemCode implements ExcelImportProblem {
  /** A required cell was empty. */
  REQUIRED_FIELD(1, "Required field"),

  /** Text was shorter than the annotated minimum. */
  MIN_LENGTH(2, "Minimum length is %d"),

  /** Text exceeded the annotated maximum. */
  MAX_LENGTH(3, "Maximum length is %d"),

  /** Text did not match the annotated regular expression. */
  PATTERN_MISMATCH(4, "Pattern mismatch: %s"),

  /** Annotation supplied an invalid Java regular expression. */
  INVALID_PATTERN_REGEX(5, "Invalid pattern regex: %s"),

  /** Text could not be converted to its target type. */
  INVALID_FORMAT(6, "Cannot convert value to %s"),

  /** Reflection could not set a DTO field. */
  MAPPING_ERROR(7, "Mapping error: %s"),

  /** An unanticipated error interrupted processing of a row. */
  UNEXPECTED_ERROR(8, "Unexpected error while processing row: %s"),

  /** A spreadsheet header did not match the column declared by the import schema. */
  HEADER_MISMATCH(9, "Expected header '%s' but found '%s'"),

  /** The configured header row was not present in the selected worksheet. */
  HEADER_ROW_NOT_FOUND(10, "Configured header row was not found"),

  /** A number was below the annotated minimum. */
  MIN_VALUE(11, "Minimum value is %s"),

  /** A number was above the annotated maximum. */
  MAX_VALUE(12, "Maximum value is %s"),

  /** Text was not one of the permitted annotation values. */
  VALUE_NOT_ALLOWED(13, "Value is not allowed");

  /** Stable error code exposed to callers. */
  private final int code;

  /** English message template; parameters are filled in by the importer where needed. */
  private final String message;

  ExcelImportProblemCode(int code, String message) {
    this.code = code;
    this.message = message;
  }

  /** {@inheritDoc} */
  public int getCode() {
    return code;
  }

  /** {@inheritDoc} */
  public String getMessage() {
    return message;
  }
}
