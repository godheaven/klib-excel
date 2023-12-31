package cl.kanopus.excel.writer;

import cl.kanopus.excel.writer.streaming.KSheet;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
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
    
    public enum DefaultStyle {
        DEFAULT_DATE,
        DEFAULT_DATETIME,
        DEFAULT_VALUE_NORMAL,
        DEFAULT_VALUE_NORMAL_DATE,
        DEFAULT_VALUE_NORMAL_DATETIME,
        
    }
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
        try (wb) {
            wb.write(out);
            wb.dispose(); // Dispose of temporary files backing this workbook on disk.
        }
        return out;
    }
    
    public CreationHelper getFactory() {
        return factory;
    }

    /**
     * Cell styles used for formatting sheets. These styles will be stored in
     * memory, in this way it is not necessary to create for each call.
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
