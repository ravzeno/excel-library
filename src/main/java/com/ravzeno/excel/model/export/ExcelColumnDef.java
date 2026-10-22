package com.ravzeno.excel.model.export;

import com.ravzeno.excel.annotation.ExcelColumn;
import com.ravzeno.excel.enums.ExcelDateFormat;
import com.ravzeno.excel.enums.ExcelNumberFormat;
import java.awt.Color;
import java.lang.reflect.Field;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Resolved export metadata for one annotated DTO field.
 * Instances are built from {@link ExcelColumn} and hold POI-ready values so reflection is done only once.
 *
 * @param name header text
 * @param index requested logical column position; {@code -1} means append
 * @param width column width in POI units (1/256th of a character)
 * @param headerColor optional header fill colour
 * @param bold whether the header font is bold
 * @param fontColor optional header font colour
 * @param fontSize optional header font size in points
 * @param dataColor optional data-cell fill colour
 * @param dataBold whether the data-cell font is bold
 * @param dataFontColor optional data-cell font colour
 * @param dataFontSize optional data-cell font size in points
 * @param align horizontal data-cell alignment
 * @param dateFormat Excel pattern used for date values
 * @param numberFormat Excel pattern used for numeric values
 * @param field source DTO field from which values are read
 */
public record ExcelColumnDef(
        String name,
        int index,
        int width,
        XSSFColor headerColor,
        boolean bold,
        XSSFColor fontColor,
        Integer fontSize,
        XSSFColor dataColor,
        boolean dataBold,
        XSSFColor dataFontColor,
        Integer dataFontSize,
        HorizontalAlignment align,
        ExcelDateFormat dateFormat,
        ExcelNumberFormat numberFormat,
        Field field) {
    /** Number of POI width units in one Excel character-width unit. */
    public static final int WPC = 256;
    /** Default column width when the annotation leaves width unspecified. */
    private static final int DEFAULT_WIDTH = 20 * WPC;

    /**
     * Resolves an {@link ExcelColumn} annotation into data that can be used directly by Apache POI.
     *
     * @param field DTO field to inspect
     * @return resolved definition, or {@code null} when the field is not annotated
     * @throws IllegalArgumentException when an annotated colour is not a valid {@code #RRGGBB} value
     */
    public static ExcelColumnDef fromField(Field field) {
        ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
        if (annotation == null) {
            return null;
        }
        int width = annotation.width() > 0 ? annotation.width() * WPC : DEFAULT_WIDTH;
        return new ExcelColumnDef(
                annotation.name(),
                annotation.index(),
                width,
                parseColor(annotation.headerCellColor()),
                annotation.headerFontBold(),
                parseColor(annotation.headerFontColor()),
                annotation.headerFontSize() > 0 ? annotation.headerFontSize() : null,
                parseColor(annotation.dataColor()),
                annotation.dataBold(),
                parseColor(annotation.dataFontColor()),
                annotation.dataFontSize() > 0 ? annotation.dataFontSize() : null,
                annotation.align(),
                annotation.dateFormat(),
                annotation.numberFormat(),
                field);
    }

    /** Converts an optional {@code #RRGGBB} annotation value to a POI RGB colour. */
    private static XSSFColor parseColor(String value) {
        if (value == null || "NONE".equalsIgnoreCase(value)) {
            return null;
        }
        try {
            Color color = Color.decode(value);
            return new XSSFColor(
                    new byte[] {(byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue()},
                    null);
        } catch (NumberFormatException exception) {
            throw new IllegalArgumentException("Invalid Excel color '" + value + "'", exception);
        }
    }
}
