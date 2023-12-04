package cl.kanopus.excel.writer.streaming;

import cl.kanopus.excel.writer.KanopusExcel;
import java.io.IOException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class KSheet {

    private static final int AUTO_LIMIT = 50000;
    private int indexRow = 0;
    private final SXSSFSheet sheet;
    private final KanopusExcel excel;

    public KSheet(KanopusExcel excel, SXSSFSheet sheet) {
        this.excel = excel;
        this.sheet = sheet;
    }

    public KRow createRow() {
        return new KRow(excel, sheet.createRow(indexRow++));
    }

    public KRow createRow(int rownum) {
        indexRow = rownum;
        return createRow();
    }

    public void createFreezePane(int colSplit, int rowSplit) {
        sheet.createFreezePane(colSplit, rowSplit);
    }

    public void autoFilter(int columns) {
        if (columns > 0 && indexRow > 0 && indexRow < AUTO_LIMIT) {
            sheet.setAutoFilter(new CellRangeAddress(0, indexRow - 1, 0, columns - 1));
        }
    }

    public void autoSize(int columns) {
        if (columns > 0 && indexRow < AUTO_LIMIT && !sheet.areAllRowsFlushed()) {
            for (int i = 0; i < columns; i++) {
                sheet.trackAllColumnsForAutoSizing();
                sheet.autoSizeColumn(i, false);
            }
        }
    }

    public void setMerge(int numRow, int untilRow, int numCol, int untilCol, boolean border) {
        CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
        sheet.addMergedRegion(cellMerge);
        if (border) {
            setBordersToMergedCells(sheet, cellMerge);
        }
    }

    protected void setBordersToMergedCells(SXSSFSheet sheet, CellRangeAddress rangeAddress) {
        RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
    }

}
