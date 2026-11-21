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
package cl.kanopus.excel.writer.iterator;

import cl.kanopus.common.data.Paginator;
import cl.kanopus.common.data.Searcher;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class ExcelReaderIteratorTest {

    private static final class Products extends ExcelReaderIterator<Integer> {

        private final List<Integer> data = new ArrayList<>();
        private int calls = 0;
        private final boolean nullPaginator;

        Products(int count, boolean nullPaginator) {
            for (int i = 0; i < count; i++) {
                data.add(i);
            }
            this.nullPaginator = nullPaginator;
        }

        @Override
        protected <I extends Searcher<?>> Paginator<Integer> findRecords(I searcher) {
            calls++;
            if (nullPaginator) {
                return null;
            }
            int from = Math.min(searcher.getOffset(), data.size());
            int to = Math.min(from + searcher.getLimit(), data.size());
            Paginator<Integer> paginator = new Paginator<>();
            paginator.setRecords(new ArrayList<>(data.subList(from, to)));
            paginator.setTotalRecords(data.size());
            return paginator;
        }

        @Override
        protected List<String> header() {
            return List.of("code", "name");
        }

        @Override
        protected List<Object> row(Integer value) {
            return List.of(value, "Product " + value);
        }
    }

    @Test
    void readAllPages() throws Exception {
        Products products = new Products(2500, false);
        try (InputStream in = products.read(new Searcher<Long>()); XSSFWorkbook wb = new XSSFWorkbook(in)) {
            XSSFSheet sheet = wb.getSheet("RECORDS");
            Assertions.assertEquals("CODE", sheet.getRow(0).getCell(0).getStringCellValue());
            Assertions.assertEquals("NAME", sheet.getRow(0).getCell(1).getStringCellValue());
            Assertions.assertEquals(2500, sheet.getLastRowNum());
            Assertions.assertEquals(2499, sheet.getRow(2500).getCell(0).getNumericCellValue());
            Assertions.assertEquals("Product 2499", sheet.getRow(2500).getCell(1).getStringCellValue());
            Assertions.assertTrue(sheet.getCTWorksheet().isSetAutoFilter());
        }
        Assertions.assertEquals(4, products.calls);
    }

    @Test
    void writeWithoutRecords() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new Products(0, true).write(new Searcher<Long>(), out);
        try (XSSFWorkbook wb = new XSSFWorkbook(new java.io.ByteArrayInputStream(out.toByteArray()))) {
            Assertions.assertEquals(0, wb.getSheet("RECORDS").getLastRowNum());
        }
    }

    @Test
    void recordsExcel() throws Exception {
        try (RecordsExcel excel = new RecordsExcel(10, false)) {
            excel.setEnableRecordStyle(true);
            Assertions.assertTrue(excel.isEnableRecordStyle());
            excel.createHeaders(List.of("a"));
            excel.createRecord(List.of("x"));
            excel.createRecord(List.of("y"));
            Assertions.assertEquals(2, excel.getTotalRecords());
            excel.autoFilter();
            excel.autoSize();
            try (XSSFWorkbook wb = new XSSFWorkbook(new java.io.ByteArrayInputStream(excel.generateOutput().toByteArray()))) {
                XSSFSheet sheet = wb.getSheet("RECORDS");
                Assertions.assertEquals("A1:A3", sheet.getCTWorksheet().getAutoFilter().getRef());
                Assertions.assertEquals("y", sheet.getRow(2).getCell(0).getStringCellValue());
            }
        }
    }
}
