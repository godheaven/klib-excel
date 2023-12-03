package cl.kanopus.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import org.junit.jupiter.api.Test;
import cl.kanopus.common.util.DesktopUtils;
import cl.kanopus.common.util.FileUtils;
import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;
import cl.kanopus.excel.writer.KanopusExcel;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ExcelWriterTest {

    @Test
    public void generateExcelOneMillionRecords() throws Exception {
        // 1millon --> 47seconds --> 19.03 MB --> 10.000 in memory

        KanopusExcel excel = new KanopusExcel(10000, true);
        KSheet sheet = excel.createSheet("RECORDS");

        // HEADER
        KRow header = sheet.createRow();
        header.createCell("Code", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Name", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Price", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");

        // RECORDS
        for (int i = 0; i < 1000000; i++) {
            KRow row = sheet.createRow();
            row.createCell("A123456789-" + i);
            row.createCell("name" + i);
            row.createCell(i);
        }

        sheet.createFreezePane(0, 1);
        sheet.autoFilter(header.getColumns());
        sheet.autoSize(header.getColumns());

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products.xlsx");

        DesktopUtils.open(file);

    }

}
