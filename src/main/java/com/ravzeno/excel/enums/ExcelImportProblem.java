package com.ravzeno.excel.enums;

/** Describes a stable machine-readable code and human-readable message for an import issue. */
public interface ExcelImportProblem {
    /**
     * Returns the stable numeric code suitable for clients and logs.
     *
     * @return numeric problem code
     */
    int getCode();
    /**
     * Returns the message template describing the problem.
     *
     * @return message template, which can include {@code String.format} placeholders
     */
    String getMessage();
}
