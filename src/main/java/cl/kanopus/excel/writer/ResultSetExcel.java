package cl.kanopus.excel.writer;

import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;
import java.util.Iterator;
import java.util.Map;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ResultSetExcel extends KanopusExcel {

    private static final String SHEET_RECORDS = "RECORDS";
    private boolean enableAutoFilter = true;
    private boolean enableFreezePane = true;

    public ResultSetExcel() {
    }

    public ResultSetExcel(int rowAccessWindowSize, boolean compressTmpFiles) {
        super(rowAccessWindowSize, compressTmpFiles);
    }

    public int createSheet(Iterator<Map<String, Object>> records) {
        return createSheet(SHEET_RECORDS, records);
    }

    public int createSheet(String sheetName, Iterator<Map<String, Object>> records) {
        int rows = 0;
        KSheet sheet = createSheet(SHEET_RECORDS);

        boolean headerIncluded = false;

        int columns = 0;
        while (records.hasNext()) {
            Map<String, Object> p = records.next();

            if (!headerIncluded) {
                KRow header = sheet.createRow();
                for (String title : p.keySet()) {
                    header.createCell(title, KanopusExcel.Style.TABLE_TITLE_NORMAL);
                    columns++;
                }
                headerIncluded = true;
            }

            KRow row = sheet.createRow();
            for (Object value : p.values()) {
                row.createCell(value, KanopusExcel.Style.TABLE_VALUE_NORMAL);
            }

        }

        if (enableAutoFilter) {
            sheet.autoFilter(columns);
        }

        if (enableFreezePane) {
            sheet.createFreezePane(0, 1);
        }
        return rows;
    }

    public void setEnableAutoFilter(boolean enableAutoFilter) {
        this.enableAutoFilter = enableAutoFilter;
    }

    public void setEnableFreezePane(boolean enableFreezePane) {
        this.enableFreezePane = enableFreezePane;
    }

}
