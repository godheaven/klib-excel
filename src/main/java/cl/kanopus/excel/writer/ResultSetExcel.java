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

    public static final String SHEET_RECORDS = "RECORDS";

    public void createSheet(Iterator<Map<String, Object>> records) {
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

        sheet.autoFilter(columns);
        sheet.autoSize(columns);
        sheet.createFreezePane(0, 1);

    }

}
