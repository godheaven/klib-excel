package cl.kanopus.excel.writer.streaming;

import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.writer.KanopusExcel;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.CellStyle;
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
        createCell(value, (KanopusExcel.Style) null, null);
    }

    public void createCell(Object value, KanopusExcel.Style style) {
        createCell(value, style, null);
    }

    public void createCell(Object value, KanopusExcel.Style style, String comment) {
        CellStyle cellStyle = calculateStyle(value, style);
        createCell(value, cellStyle, comment);
    }

    public void createCell(Object value, CellStyle style, String comment) {
        SXSSFCell cell = row.createCell(this.columns++);
        if (style != null) {
            cell.setCellStyle(style);
        }
        setCellValue(cell, value);

        if (!Utils.isNullOrEmpty(comment)) {
            setComment(cell, comment);
        }
    }

    private void setCellValue(SXSSFCell cell, Object value) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
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
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value == null ? "" : String.valueOf(value));
        }

    }

    private void setComment(SXSSFCell cell, String comment) {
        // Assign the comment to the cell
        Comment cellcomment = cell.getSheet().createDrawingPatriarch().createCellComment(
                new XSSFClientAnchor(0, 0, 0, 0, cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex() + 2, cell.getRowIndex() + 3));
        RichTextString richText = excel.getFactory().createRichTextString(comment);
        cellcomment.setString(richText);
    }

    private CellStyle calculateStyle(Object value, KanopusExcel.Style style) {
        CellStyle cellStyle = null;
        if (style != null) {
            cellStyle = excel.styles.get(style.name());
        }

        if (value instanceof Date || value instanceof LocalDate) {
            if (style == null) {
                cellStyle = excel.styles.get(KanopusExcel.DefaultStyle.DEFAULT_DATE.name());
            } else if (style == KanopusExcel.Style.TABLE_VALUE_NORMAL) {
                cellStyle = excel.styles.get(KanopusExcel.DefaultStyle.DEFAULT_VALUE_NORMAL_DATE.name());
            }
        } else if (value instanceof LocalDateTime) {
            if (style == null) {
                cellStyle = excel.styles.get(KanopusExcel.DefaultStyle.DEFAULT_DATETIME.name());
            } else if (style == KanopusExcel.Style.TABLE_VALUE_NORMAL) {
                cellStyle = excel.styles.get(KanopusExcel.DefaultStyle.DEFAULT_VALUE_NORMAL_DATETIME.name());
            }
        }

        return cellStyle;
    }

}
