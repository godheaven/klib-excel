package cl.kanopus.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import org.junit.jupiter.api.Test;
import cl.kanopus.common.util.DesktopUtils;
import cl.kanopus.common.util.FileUtils;
import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.writer.KanopusExcel;
import cl.kanopus.excel.writer.ResultSetExcel;
import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ExcelWriterTest {
  
    @Test
    public void generateResultSetExcelOneMillion() throws Exception {
        // 1millon --> 11seconds --> 23 MB --> 1000 in memory

        Iterator records = new Iterator<Map<String, Object>>() {
            private int index = 0;

            @Override
            public boolean hasNext() {
                return index < 1000;
            }

            @Override
            public Map<String, Object> next() {
                Map<String, Object> rs = new LinkedHashMap<>();
                rs.put("Title1", "A123456789-" + index);
                rs.put("Title2", index);
                rs.put("Title3", new Date());
                rs.put("Title4", LocalDate.now());
                rs.put("Title5", LocalDateTime.now());
                index++;

                return rs;
            }
        };

        ResultSetExcel excel = new ResultSetExcel(1000, true);
        excel.createSheet(records);

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products_rs_test.xlsx");
        DesktopUtils.open(file);
    }
    
    
    @Test
    public void generateExcelOneMillionRecords() throws Exception {
        //1millon --> 11seconds --> 25 MB --> 10.000 in memory

        KanopusExcel excel = new KanopusExcel(1000, true);
        excel.setAutoformat(true);
        
        KSheet sheet = excel.createSheet("RECORDS");

        // HEADER
        KRow header = sheet.createRow();
        header.createCell("Text", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Integer", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Date", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("LocalDate", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("LocalDatetime", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");

        // RECORDS
        for (int i = 0; i < 100; i++) {
            KRow row = sheet.createRow();
            row.createCell("A123456789-" + i);
            row.createCell(i);
            row.createCell(new Date());
            row.createCell(LocalDate.now());
            row.createCell(LocalDateTime.now());
        }

        sheet.createFreezePane(0, 1);
        sheet.autoFilter(header.getColumns());
        sheet.autoSize(header.getColumns());

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products.xlsx");

        DesktopUtils.open(file);

    }

}
