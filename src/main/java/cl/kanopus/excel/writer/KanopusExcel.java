package cl.kanopus.excel.writer;

import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.writer.streaming.KSheet;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class KanopusExcel {

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

    public final Map<String, CellStyle> styles = new HashMap<>();

    private final SXSSFWorkbook wb;

    private final CreationHelper factory;
    private boolean autoformat = false;

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
        try (wb) {
            wb.write(out);
            wb.dispose(); // Dispose of temporary files backing this workbook on disk.
        }
        return out;
    }

    public CreationHelper getFactory() {
        return factory;
    }

    public boolean isAutoformat() {
        return autoformat;
    }

    public void setAutoformat(boolean autoformat) {
        this.autoformat = autoformat;
    }

    public String applyAutoFormatDate(Date source) {
        return Utils.getDateFormat((Date) source);
    }

    public String applyAutoFormatDate(LocalDate source) {
        return Utils.getDateFormat(source);
    }

    public String applyAutoFormatDate(LocalDateTime source) {
        return Utils.getDateTimeFormat(source);
    }

    public String applyAutoFormatBoolean(Boolean source) {
        return source ? "YES" : "NO";
    }

    /**
     * Cell styles used for formatting sheets. These styles will be stored in
     * memory, in this way it is not necessary to create for each call.
     */
    private void createStyles(SXSSFWorkbook workbook) {
        if (styles.isEmpty()) {

            /* Style TABLE_TITLE_NORMAL */
            styles.put(Style.TABLE_TITLE_NORMAL.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex()));
            /* Style TABLE_TITLE_INFO */
            styles.put(Style.TABLE_TITLE_INFO.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getIndex()));
            /* Style TABLE_TITLE_REQUIRED */
            styles.put(Style.TABLE_TITLE_REQUIRED.name(), createStyle(workbook, true, Font.COLOR_RED, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));
            /* Style TABLE_TITLE_OPTIONAL */
            styles.put(Style.TABLE_TITLE_OPTIONAL.name(), createStyle(workbook, true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));

            /* Style TABLE_VALUE_NORMAL */
            styles.put(Style.TABLE_VALUE_NORMAL.name(), createStyle(workbook, false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.WHITE.getIndex()));
            /* Style TABLE_VALUE_INSERT */
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
            font.setBold(true);
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
        return style;
    }

}
