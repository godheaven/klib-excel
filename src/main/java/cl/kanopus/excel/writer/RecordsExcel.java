package cl.kanopus.excel.writer;

import cl.kanopus.excel.writer.streaming.KSheet;
import cl.kanopus.excel.writer.streaming.KRow;
import java.util.List;


public class RecordsExcel extends KanopusExcel {

    public static final String SHEET_RECORDS = "RECORDS";

    private int columns = 0;
    private int totalRecords = 0;
    private boolean enableRecordStyle = false;
    private KSheet sheet;

    public RecordsExcel() {
    }

    public RecordsExcel(int rowAccessWindowSize, boolean compressTmpFiles) {
        super(rowAccessWindowSize, compressTmpFiles);
        this.sheet = createSheet(SHEET_RECORDS);
    }

    public void createHeaders(List<String> titles) {
        KRow header = sheet.createRow();
        for (String title : titles) {
            header.createCell(title.toUpperCase(), Style.TABLE_TITLE_NORMAL);
        }
        sheet.createFreezePane(0, 1);
        columns = titles.size();
    }

    public void createRecord(List<Object> values) {
        KRow row = sheet.createRow();
        for (Object value : values) {
            row.createCell(value, enableRecordStyle ? Style.TABLE_VALUE_NORMAL : null);
        }
        this.totalRecords++;
    }

    public void autoSize() {
         sheet.autoFilter(columns);
    }

    public void autoFilter() {
        sheet.autoSize(columns);
    }

    public boolean isEnableRecordStyle() {
        return enableRecordStyle;
    }

    public void setEnableRecordStyle(boolean enableRecordStyle) {
        this.enableRecordStyle = enableRecordStyle;
    }

    public int getTotalRecords() {
        return totalRecords;
    }

}
