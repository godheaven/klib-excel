package cl.kanopus.excel.writer;

import cl.kanopus.common.util.Utils;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 *
 * @author pablo
 */
public abstract class AbstractExcel {

    public enum Format {

        MONEY("$#,##0"),
        DATE("dd/MM/yyyy");

        private final String text;

        Format(String text) {
            this.text = text;
        }

        public String getText() {
            return text;
        }

    }

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

    private Map<String, HSSFCellStyle> styles;

    protected HSSFWorkbook wb = null;
    protected HSSFRow row = null;
    private final HSSFCreationHelper factory;

    public AbstractExcel() {
        this.wb = new HSSFWorkbook();
        this.factory = this.wb.getCreationHelper();
        createStyles();
    }

    public final ByteArrayOutputStream generateOutput() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wb.write(out);
        // Dispose of temporary files backing this workbook on disk.
        //Calling this method will render the workbook unusable.
        wb.close();
        return out;
    }

    protected void createCell(int i, Object value) {
        createCell(i, value, null, null, null);
    }

    protected void createCell(int i, Object value, Style style) {
        createCell(i, value, style, null, null);
    }

    protected void createCell(int i, Object value, Style style, Format format) {
        createCell(i, value, style, format, null);
    }

    protected void createCell(int i, Object value, Style style, String comment) {
        createCell(i, value, style, null, comment);
    }

    protected void createCell(int i, Object value, Style style, Format format, String comment) {

        HSSFCell cell = row.createCell(i);

        //1. Create the date cell style
        if (style != null) {
            cell.setCellStyle(styles.get(style.name()));
        }

        if (format != null && format != Format.DATE) {
            cell.getCellStyle().setDataFormat(factory.createDataFormat().getFormat(format.getText()));
        }

        if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Long && format == Format.MONEY) {
            cell.setCellValue((Double) (value == null ? null : ((Long) value).doubleValue()));
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) (value == null ? null : (Calendar) value));
        } else if (value instanceof Date) {
            cell.setCellValue(Utils.getDateFormat((Date) value, Format.DATE.getText()));
        } else if (value instanceof Boolean) {
            cell.setCellValue(((Boolean) value) ? "SI" : "NO");
        } else {
            cell.setCellValue(value == null ? "" : String.valueOf(value));
        }

        if (comment != null && !comment.trim().isEmpty()) {
            RichTextString richText = factory.createRichTextString(comment);
            Drawing drawing = row.getSheet().createDrawingPatriarch();
            // When the comment box is visible, have it show in a 1x3 space
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
            Comment cellComment = drawing.createCellComment(anchor);
            cellComment.setString(richText);
            //cellComment.setAuthor(author);

            // Set the row and column here
            cellComment.setRow(cell.getRowIndex());
            cellComment.setColumn(cell.getColumnIndex());

            // Assign the comment to the cell
            cell.setCellComment(cellComment);
        }

    }

    protected void setMerge(HSSFSheet sheet, int numRow, int untilRow, int numCol, int untilCol, boolean border) {
        CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
        sheet.addMergedRegion(cellMerge);
        if (border) {
            setBordersToMergedCells(sheet, cellMerge);
        }

    }

    protected void setBordersToMergedCells(HSSFSheet sheet, CellRangeAddress rangeAddress) {
        RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
    }

    /**
     * Cell styles used for formatting sheets. These styles will be stored in
     * memory, in this way it is not necessary to create for each call.
     */
    private Map<String, HSSFCellStyle> createStyles() {
        styles = new HashMap<>();

        /* Style TABLE_TITLE_NORMAL */
        styles.put(Style.TABLE_TITLE_NORMAL.name(), createStyle(true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex()));
        /* Style TABLE_TITLE_INFO */
        styles.put(Style.TABLE_TITLE_INFO.name(), createStyle(true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getIndex()));
        /* Style TABLE_TITLE_REQUIRED */
        styles.put(Style.TABLE_TITLE_REQUIRED.name(), createStyle(true, Font.COLOR_RED, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));
        /* Style TABLE_TITLE_OPTIONAL */
        styles.put(Style.TABLE_TITLE_OPTIONAL.name(), createStyle(true, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.YELLOW.getIndex()));

        /* Style TABLE_VALUE_NORMAL */
        styles.put(Style.TABLE_VALUE_NORMAL.name(), createStyle(false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.WHITE.getIndex()));
        /* Style TABLE_VALUE_INSERT */
        styles.put(Style.TABLE_VALUE_INSERT.name(), createStyle(false, Font.COLOR_NORMAL, HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex()));

        /* Style FONT_BLACK_BOLD */
        HSSFCellStyle style;
        Font font;

        font = wb.createFont();
        font.setBold(true);
        style = wb.createCellStyle();
        style.setFont(font);
        styles.put(Style.FONT_BLACK_BOLD.name(), style);

        /* Style FONT_BLACK_NORMAL */
        font = wb.createFont();
        font.setBold(true);
        style = wb.createCellStyle();
        style.setFont(font);
        styles.put(Style.FONT_BLACK_NORMAL.name(), style);

        /* Style FONT_RED_BOLD */
        font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.RED.getIndex());
        style = wb.createCellStyle();
        style.setFont(font);
        styles.put(Style.FONT_RED_BOLD.name(), style);

        return styles;
    }

    private HSSFCellStyle createStyle(boolean bold, short fontColor, short backgroundColor) {
        Font font = wb.createFont();
        font.setBold(bold);
        font.setColor(fontColor);
        HSSFCellStyle style = wb.createCellStyle();
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
