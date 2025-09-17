/*-
 * !--
 * For support and inquiries regarding this library, please contact:
 *   soporte@kanopus.cl
 *
 * Project website:
 *   https://www.kanopus.cl
 * %%
 * Copyright (C) 2025 Pablo Díaz Saavedra
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * --!
 */
package cl.kanopus.excel;

import cl.kanopus.excel.reader.ExcelReader;
import cl.kanopus.excel.reader.RowProcess;
import cl.kanopus.excel.reader.SheetEventListener;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.net.URL;
import java.util.HashMap;


class ExcelReaderTest {

    @Test
    void readExcel() throws Exception {

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

    @Test
    void readExcelXlsx() throws Exception {

        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("RECORDS", new RowProcess() {
            @Override
            public void process(HashMap<String, String> row) throws Exception {
                System.out.println("row-->" + row);
            }
        }));

        URL resource = ExcelReaderTest.class.getClassLoader().getResource("test_reader.xlsx");
        reader.processAllSheets(new File(resource.getFile()));

        Assertions.assertEquals(1, reader.getTotalProcessedSheets());
        Assertions.assertEquals(9, reader.getTotalProcessedRecords());
    }
}
