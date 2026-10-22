package com.ravzeno.excel.model.export;

/**
 * In-memory Excel export payload ready for an HTTP response or file write.
 * For a large response, prefer the service {@code writeTo(..., OutputStream)} API instead.
 *
 * @param bytes complete XLSX binary content
 * @param filename suggested output filename
 * @param contentType HTTP media type; blank values become the standard XLSX type
 */
public record ExcelExportResult(byte[] bytes, String filename, String contentType) {
    /** Standard HTTP media type for XLSX content. */
    public static final String XLSX_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /** Normalizes the optional HTTP content type. */
    public ExcelExportResult {
        contentType = contentType == null || contentType.isBlank() ? XLSX_CONTENT_TYPE : contentType;
    }

    /**
     * Creates a result using the standard XLSX HTTP media type.
     *
     * @param bytes complete XLSX binary content
     * @param filename suggested output filename
     */
    public ExcelExportResult(byte[] bytes, String filename) {
        this(bytes, filename, XLSX_CONTENT_TYPE);
    }
}
