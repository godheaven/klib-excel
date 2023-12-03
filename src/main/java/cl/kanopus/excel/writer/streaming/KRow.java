package cl.kanopus.excel.writer.streaming;

import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.writer.KanopusExcel;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

public class KRow {

    private final SXSSFRow row;
    private final KanopusExcel excel;
    private int columns = 0;

    public KRow(KanopusExcel excel, SXSSFRow row) {
        this.excel = excel;
        this.row = row;
    }

    public int getColumns() {
        return columns;
    }

    public void createCell(Object value) {
        createCell(value, null, null);
    }

    public void createCell(Object value, KanopusExcel.Style style) {
        createCell(value, style, null);
    }

    public void createCell(Object value, KanopusExcel.Style style, String comment) {

        SXSSFCell cell = row.createCell(this.columns++);

        // 1. Create the date cell style
        if (style != null) {
            cell.setCellStyle(KanopusExcel.styles.get(style.name()));
        }

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((value == Boolean.TRUE) ? "YES" : "NO");
        } else {
            cell.setCellValue(value == null ? "" : String.valueOf(value));
        }

        if (!Utils.isNullOrEmpty(comment)) {
            // Assign the comment to the cell
            Comment cellcomment = cell.getSheet().createDrawingPatriarch().createCellComment(
                    new XSSFClientAnchor(0, 0, 0, 0, cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex() + 2, cell.getRowIndex() + 3));
            RichTextString richText = excel.getFactory().createRichTextString(comment);
            cellcomment.setString(richText);
        }

    }

}