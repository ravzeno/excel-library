package com.ravzeno.excel.enums;

/** Expected type for an imported column. */
public enum ExcelDataType {
    /** Infer conversion from the target Java field type. */
    AUTO,

    /** Keep the cell value as text. */
    STRING,

    /** Convert the cell value to a number. */
    NUMERIC,

    /** Convert the cell value to a date. */
    DATE
}
