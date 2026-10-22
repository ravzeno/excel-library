package com.ravzeno.excel.model.import_;

import com.ravzeno.excel.enums.ExcelImportProblem;

/**
 * Validation or mapping issue found in one imported cell.
 * Row and column indexes are zero-based physical spreadsheet coordinates.
 *
 * @param row zero-based physical row index
 * @param columnIndex zero-based physical column index
 * @param columnName annotated target field name, when known
 * @param value formatted source cell text
 * @param problemCode stable error classification
 * @param message resolved human-readable error detail
 */
public record ExcelImportError(int row, int columnIndex, String columnName, String value,
        ExcelImportProblem problemCode, String message) {}
