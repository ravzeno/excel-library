package com.ravzeno.excel.model.import_;

import java.util.List;

/**
 * Outcome of importing a worksheet. Valid rows and row-level errors are returned together.
 *
 * @param data successfully mapped DTO rows; copied defensively
 * @param errors cell and mapping errors; copied defensively
 * @param totalRows number of non-empty rows considered for import
 * @param validRows number of rows added to {@code data}
 * @param invalidRows number of rows rejected because of one or more errors
 * @param <T> DTO type requested by the caller
 */
public record ExcelImportResult<T>(List<T> data, List<ExcelImportError> errors, int totalRows,
        int validRows, int invalidRows) {
    /** Protects both result collections from caller mutation. */
    public ExcelImportResult { data = List.copyOf(data); errors = List.copyOf(errors); }
}
