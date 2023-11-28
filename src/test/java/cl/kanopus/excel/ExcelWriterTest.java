package cl.kanopus.excel;

import cl.kanopus.common.util.Utils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/**
 *
 * @author Pablo Diaz Saavedra
 * @email pabloandres.diazsaavedra@gmail.com
 */
public class ExcelWriterTest {

    public ExcelWriterTest() {
    }

    @Test
    public void testIsNullOrEmpty() throws Exception {

        Assertions.assertTrue(Utils.isNullOrEmpty(""));
        Assertions.assertTrue(Utils.isNullOrEmpty("   "));

    }

}
