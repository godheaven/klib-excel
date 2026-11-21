/*-
 * !--
 * For support and inquiries regarding this library, please contact:
 *   soporte@kanopus.cl
 *
 * Project website:
 *   https://www.kanopus.cl
 * %%
 * Copyright (C) 2025 Pablo Díaz Saavedra
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * --!
 */
package cl.kanopus.excel.writer;

import cl.kanopus.excel.writer.streaming.KSheet;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Streaming (SXSSF) XLSX writer. An instance is single-use: {@link #write(OutputStream)} / {@link #generateOutput()} write the workbook and release its temporary files. Use try-with-resources (or {@link #close()}) so the temporary files are also removed when the workbook is never written.
 */
public class KanopusExcel implements AutoCloseable {

    public enum Style {
        FONT_BLACK_BOLD,
        FONT_BLACK_NORMAL,
        FONT_RED_BOLD,
        TABLE_TITLE_NORMAL,
        TABLE_TITLE_INFO,
        TABLE_TITLE_REQUIRED,
        TABLE_TITLE_OPTIONAL,
        TABLE_VALUE_NORMAL,
        TABLE_VALUE_INSERT
    }

    public enum DefaultStyle {
        DEFAULT_DATE,
        DEFAULT_DATETIME,
        DEFAULT_VALUE_NORMAL,
        DEFAULT_VALUE_NORMAL_DATE,
        DEFAULT_VALUE_NORMAL_DATETIME,
    }

    /**
     * @deprecated use {@link #getStyle(Style)}, {@link #getStyle(DefaultStyle)}, {@link #getStyle(String)} or {@link #registerStyle(String, CellStyle)}.
     */
    @Deprecated(since = "4.20.0", forRemoval = true)
    public final Map<String, CellStyle> styles = new HashMap<>();

    private final SXSSFWorkbook wb;

    private final CreationHelper factory;

    public KanopusExcel() {
        this.wb = new SXSSFWorkbook(new XSSFWorkbook(), 5000, true, false);
        this.factory = this.wb.getCreationHelper();
        createStyles(this.wb);
    }

    public KanopusExcel(int rowAccessWindowSize, boolean compressTmpFiles) {
        this.wb = new SXSSFWorkbook(new XSSFWorkbook(), rowAccessWindowSize, compressTmpFiles, false);
        this.factory = this.wb.getCreationHelper();
        createStyles(this.wb);
    }

    public KSheet createSheet(String sheetname) {
        return new KSheet(this, wb.createSheet(sheetname));
    }

    public final ByteArrayOutputStream generateOutput() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        write(out);
        return out;
    }

    /**
     * Writes the workbook to the given stream (which is not closed) and releases the workbook resources, including its temporary files, even on failure.
     */
    public final void write(OutputStream out) throws IOException {
        try {
            wb.write(out);
        } finally {
            close();
        }
    }

    /** Disposes the temporary files backing this workbook and closes it. Safe to call more than once. */
    @Override
    public void close() throws IOException {
        try {
            wb.dispose(); // Dispose of temporary files backing this workbook on disk.
        } finally {
            wb.close();
        }
    }

    public CellStyle getStyle(Style style) {
        return style == null ? null : getStyle(style.name());
    }

    public CellStyle getStyle(DefaultStyle style) {
        return style == null ? null : getStyle(style.name());
    }

    /** Returns a predefined style (by {@link Style}/{@link DefaultStyle} name) or a style added with {@link #registerStyle(String, CellStyle)}. */
    public CellStyle getStyle(String name) {
        return styles.get(name);
    }

    /** Registers a custom style created with {@link #createCellStyle()}, retrievable with {@link #getStyle(String)}. */
    public void registerStyle(String name, CellStyle style) {
        styles.put(name, style);
    }

    /** Creates a new cell style in this workbook (to customize and pass to {@code KRow.createCell(Object, CellStyle, String)}). */
    public CellStyle createCellStyle() {
        return wb.createCellStyle();
    }

    /** Creates a new font in this workbook. */
    public Font createFont() {
        return wb.createFont();
    }

    public CreationHelper getFactory() {
        return factory;
    }

    /**
     * Cell styles used for formatting sheets. These styles will be stored in memory, in this way it is not necessary to create for each call.
     */
    private void createStyles(SXSSFWorkbook workbook) {
        if (styles.isEmpty()) {

            short dateFormat = workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yyyy");
            short dateTimeFormat = workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yyyy HH:mm");

            /* Style TABLE_TITLE */
            styles.put(Style.TABLE_TITLE_NORMAL.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex()));
            styles.put(Style.TABLE_TITLE_INFO.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getIndex()));
            styles.put(Style.TABLE_TITLE_REQUIRED.name(), createStyle(workbook, true, Font.COLOR_RED, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));
            styles.put(Style.TABLE_TITLE_OPTIONAL.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));

            /* Style TABLE_VALUE */
            styles.put(DefaultStyle.DEFAULT_DATE.name(), createStyle(workbook, dateFormat));
            styles.put(DefaultStyle.DEFAULT_DATETIME.name(), createStyle(workbook, dateTimeFormat));

            /* VALUE NORMAL */
            styles.put(Style.TABLE_VALUE_NORMAL.name(), createStyle(workbook, false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.WHITE.getIndex()));
            styles.put(DefaultStyle.DEFAULT_VALUE_NORMAL_DATE.name(), createStyle(workbook, false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.WHITE.getIndex(), dateFormat));
            styles.put(DefaultStyle.DEFAULT_VALUE_NORMAL_DATETIME.name(), createStyle(workbook, false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.WHITE.getIndex(), dateTimeFormat));

            /* VALUE INSERT */
            styles.put(Style.TABLE_VALUE_INSERT.name(), createStyle(workbook, false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex()));

            /* Style FONT_BLACK_BOLD */
            CellStyle style;
            Font font;

            font = workbook.createFont();
            font.setBold(true);
            style = workbook.createCellStyle();
            style.setFont(font);
            styles.put(Style.FONT_BLACK_BOLD.name(), style);

            /* Style FONT_BLACK_NORMAL */
            font = workbook.createFont();
            font.setBold(false);
            style = workbook.createCellStyle();
            style.setFont(font);
            styles.put(Style.FONT_BLACK_NORMAL.name(), style);

            /* Style FONT_RED_BOLD */
            font = workbook.createFont();
            font.setBold(true);
            font.setColor(IndexedColors.RED.getIndex());
            style = workbook.createCellStyle();
            style.setFont(font);
            styles.put(Style.FONT_RED_BOLD.name(), style);
        }
    }

    private CellStyle createStyle(SXSSFWorkbook workbook, boolean bold, short fontColor, short backgroundColor) {
        return createStyle(workbook, bold, fontColor, backgroundColor, null);
    }

    private CellStyle createStyle(SXSSFWorkbook workbook, boolean bold, short fontColor, short backgroundColor, Short format) {
        Font font = workbook.createFont();
        font.setBold(bold);
        font.setColor(fontColor);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(backgroundColor); // Background Colors
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        if (format != null) {
            style.setDataFormat(format);
        }
        return style;
    }

    private CellStyle createStyle(SXSSFWorkbook workbook, short format) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(format);
        return style;
    }
}
