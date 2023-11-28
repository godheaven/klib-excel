package cl.kanopus.excel;

import cl.kanopus.common.util.DesktopUtils;
import cl.kanopus.common.util.FileUtils;
import cl.kanopus.excel.writer.ProductTO;
import cl.kanopus.excel.writer.ProductsExcel;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import org.junit.jupiter.api.Test;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ExcelWriterTest {

    @Test
    public void testGenerateExcel() throws Exception {

        ArrayList<ProductTO> products = new ArrayList<>();
        products.add(new ProductTO("A123456789", "ASPIRINA", 1000));
        products.add(new ProductTO("B123456789", "BETAMETASONA", 1000));
        products.add(new ProductTO("C123456789", "CLORFENAMINA", 1000));
        products.add(new ProductTO("D123456789", "DERMOCREAM", 1000));
        products.add(new ProductTO("E123456789", "EUCALIPTUS", 1000));
        
        ProductsExcel excel = new ProductsExcel();
        excel.createSheetProducts(products.iterator());

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products.xlsx");
 
        DesktopUtils.open(file);

    }

}
