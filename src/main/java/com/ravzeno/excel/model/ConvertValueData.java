package com.ravzeno.excel.model;

/**
 * Internal result of converting one raw spreadsheet value.
 *
 * @param value converted value; {@code null} when conversion failed or input is empty
 * @param isConvertSuccess whether the conversion completed successfully
 */
public record ConvertValueData(Object value, boolean isConvertSuccess) {
    /**
     * Creates the canonical failed conversion result.
     *
     * @return result containing no value and a failed status
     */
    public static ConvertValueData failure() {
        return new ConvertValueData(null, false);
    }
}
