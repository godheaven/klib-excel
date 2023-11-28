package cl.kanopus.excel;

import cl.kanopus.excel.reader.ExcelReader;
import cl.kanopus.excel.reader.RowProcess;
import cl.kanopus.excel.reader.SheetEventListener;
import java.io.File;
import java.net.URL;
import java.util.HashMap;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ExcelReaderTest {

    @Test
    public void readExcel() throws Exception {

        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("RECORDS", new RowProcess() {
            @Override
            public void process(HashMap<String, String> row) throws Exception {
                System.out.println("row-->" + row);
            }
        }));

        URL resource = ExcelReaderTest.class.getClassLoader().getResource("test_reader.xls");
        reader.processAllSheets(new File(resource.getFile()));

        Assertions.assertEquals(1, reader.getTotalProcessedSheets());
        Assertions.assertEquals(9, reader.getTotalProcessedRecords());
    }

}
