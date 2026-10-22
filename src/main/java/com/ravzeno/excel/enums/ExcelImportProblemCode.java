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
    UNEXPECTED_ERROR(8, "Unexpected error while processing row: %s");

    /** Stable error code exposed to callers. */
    private final int code;
    /** English message template; parameters are filled in by the importer where needed. */
    private final String message;
    ExcelImportProblemCode(int code, String message) { this.code = code; this.message = message; }
    /** {@inheritDoc} */
    public int getCode() { return code; }
    /** {@inheritDoc} */
    public String getMessage() { return message; }
}
