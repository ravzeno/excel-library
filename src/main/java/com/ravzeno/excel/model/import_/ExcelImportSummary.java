package com.ravzeno.excel.model.import_;

import java.util.List;

/** Counts and validation feedback returned by streaming import operations. */
public record ExcelImportSummary(
    int totalRows, int validRows, int invalidRows, List<ExcelImportError> errors) {
  public ExcelImportSummary {
    errors = List.copyOf(errors);
  }
}
