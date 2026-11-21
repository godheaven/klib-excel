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

import cl.kanopus.excel.writer.KanopusExcel;
import cl.kanopus.excel.writer.ResultSetExcel;
import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class ExcelWriterTest {

    private static final Path OUT = Path.of("target", "excel-writer-test");

    @BeforeAll
    static void setUp() throws Exception {
        Files.createDirectories(OUT);
    }

    private static XSSFWorkbook readBack(ByteArrayOutputStream output, String fileName) throws Exception {
        Files.write(OUT.resolve(fileName), output.toByteArray());
        return new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
    }

    private static Iterator<Map<String, Object>> records(int count) {
        return new Iterator<>() {
            private int index = 0;

            @Override
            public boolean hasNext() {
                return index < count;
            }

            @Override
            public Map<String, Object> next() {
                Map<String, Object> rs = new LinkedHashMap<>();
                rs.put("Title1 String", "A123456789-" + index);
                rs.put("Title2 int", index);
                rs.put("Title3 Date", new Date());
                rs.put("Title4 LocalDate", LocalDate.now());
                rs.put("Title5 LocalDateTime", LocalDateTime.now());
                index++;
                return rs;
            }
        };
    }

    @Test
    void generateResultSetExcel() throws Exception {
        ResultSetExcel excel = new ResultSetExcel(100, true);
        int rows = excel.createSheet(records(1000));
        Assertions.assertEquals(1000, rows);

        try (XSSFWorkbook wb = readBack(excel.generateOutput(), "products_rs_test.xlsx")) {
            XSSFSheet sheet = wb.getSheet("RECORDS");
            Assertions.assertNotNull(sheet);
            Assertions.assertEquals(1000, sheet.getLastRowNum());
            Assertions.assertEquals("Title1 String", sheet.getRow(0).getCell(0).getStringCellValue());
            Assertions.assertEquals("A123456789-999", sheet.getRow(1000).getCell(0).getStringCellValue());
            Assertions.assertEquals(999, sheet.getRow(1000).getCell(1).getNumericCellValue());
            Assertions.assertTrue(DateUtil.isCellDateFormatted(sheet.getRow(1).getCell(2)));
            Assertions.assertTrue(sheet.getCTWorksheet().isSetAutoFilter());
            Assertions.assertNotNull(sheet.getPaneInformation());
        }
    }

    @Test
    void resultSetExcelUsesSheetNameAndStableColumns() throws Exception {
        Map<String, Object> first = new LinkedHashMap<>();
        first.put("B", "b1");
        first.put("A", 1);
        first.put("C", "c1");
        Map<String, Object> second = new HashMap<>();
        second.put("C", "c2");
        second.put("A", 2);
        second.put("EXTRA", "ignored");

        ResultSetExcel excel = new ResultSetExcel();
        excel.setEnableAutoFilter(false);
        excel.setEnableFreezePane(false);
        Assertions.assertEquals(2, excel.createSheet("Customers", List.of(first, second).iterator()));
        Assertions.assertEquals(0, excel.createSheet("Empty", List.<Map<String, Object>>of().iterator()));

        try (XSSFWorkbook wb = readBack(excel.generateOutput(), "resultset-columns.xlsx")) {
            Assertions.assertNull(wb.getSheet("RECORDS"));
            Assertions.assertNotNull(wb.getSheet("Empty"));
            XSSFSheet sheet = wb.getSheet("Customers");
            Assertions.assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
            Assertions.assertNull(sheet.getPaneInformation());
            XSSFRow header = sheet.getRow(0);
            Assertions.assertEquals(3, header.getLastCellNum());
            Assertions.assertEquals("B", header.getCell(0).getStringCellValue());
            Assertions.assertEquals("A", header.getCell(1).getStringCellValue());
            Assertions.assertEquals("C", header.getCell(2).getStringCellValue());
            XSSFRow row = sheet.getRow(2);
            Assertions.assertEquals("", row.getCell(0).getStringCellValue());
            Assertions.assertEquals(2, row.getCell(1).getNumericCellValue());
            Assertions.assertEquals("c2", row.getCell(2).getStringCellValue());
            Assertions.assertEquals(3, row.getLastCellNum());
        }
    }

    @Test
    void writesTypedCells() throws Exception {
        Calendar calendar = new GregorianCalendar(2024, Calendar.MARCH, 5);
        Timestamp timestamp = Timestamp.valueOf(LocalDateTime.of(2024, 3, 5, 10, 30));
        UUID uuid = UUID.randomUUID();
        List<Object> values = List.of("text", 1, 2L, 3.5d, new BigDecimal("1234.56"), BigInteger.valueOf(99), 1.5f, (short) 7, true, new Date(), calendar, LocalDate.of(2024, 3, 5),
                LocalDateTime.of(2024, 3, 5, 10, 30), timestamp, uuid);

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (KanopusExcel excel = new KanopusExcel()) {
            KSheet sheet = excel.createSheet("TYPES");
            KRow row = sheet.createRow();
            for (Object value : values) {
                row.createCell(value);
            }
            row.createCell(null, KanopusExcel.Style.TABLE_VALUE_NORMAL);
            KRow styled = sheet.createRow(3);
            styled.createCell(timestamp, KanopusExcel.Style.TABLE_VALUE_NORMAL);
            styled.createCell(calendar, KanopusExcel.Style.TABLE_VALUE_NORMAL);
            styled.createCell("with comment", KanopusExcel.Style.FONT_RED_BOLD, "a comment");
            Assertions.assertEquals(values.size() + 1, row.getColumns());
            excel.write(output);
        }

        try (XSSFWorkbook wb = readBack(output, "types.xlsx")) {
            XSSFRow row = wb.getSheet("TYPES").getRow(0);
            Assertions.assertEquals("text", row.getCell(0).getStringCellValue());
            double[] numbers = {1, 2, 3.5, 1234.56, 99, 1.5, 7};
            for (int i = 0; i < numbers.length; i++) {
                Cell cell = row.getCell(i + 1);
                Assertions.assertEquals(CellType.NUMERIC, cell.getCellType(), "column " + (i + 1));
                Assertions.assertEquals(numbers[i], cell.getNumericCellValue(), 0.0001);
                Assertions.assertFalse(DateUtil.isCellDateFormatted(cell));
            }
            Assertions.assertTrue(row.getCell(8).getBooleanCellValue());
            for (int i = 9; i <= 13; i++) {
                Assertions.assertTrue(DateUtil.isCellDateFormatted(row.getCell(i)), "column " + i);
            }
            Assertions.assertEquals(LocalDateTime.of(2024, 3, 5, 0, 0), row.getCell(10).getLocalDateTimeCellValue());
            Assertions.assertFalse(row.getCell(10).getCellStyle().getDataFormatString().contains("HH"));
            Assertions.assertTrue(row.getCell(12).getCellStyle().getDataFormatString().contains("HH:mm"));
            Assertions.assertTrue(row.getCell(13).getCellStyle().getDataFormatString().contains("HH:mm"));
            Assertions.assertEquals(LocalDateTime.of(2024, 3, 5, 10, 30), row.getCell(13).getLocalDateTimeCellValue());
            Assertions.assertEquals(uuid.toString(), row.getCell(14).getStringCellValue());
            Assertions.assertEquals("", row.getCell(15).getStringCellValue());

            XSSFRow styled = wb.getSheet("TYPES").getRow(3);
            Assertions.assertTrue(styled.getCell(0).getCellStyle().getDataFormatString().contains("HH:mm"));
            Assertions.assertTrue(DateUtil.isCellDateFormatted(styled.getCell(1)));
            Assertions.assertEquals("a comment", styled.getCell(2).getCellComment().getString().getString());
        }
    }

    @Test
    @SuppressWarnings("removal")
    void stylesAreEncapsulated() throws Exception {
        try (KanopusExcel excel = new KanopusExcel(10, false)) {
            Assertions.assertNotNull(excel.getFactory());
            Assertions.assertTrue(fontOf(excel, KanopusExcel.Style.FONT_BLACK_BOLD).getBold());
            Assertions.assertFalse(fontOf(excel, KanopusExcel.Style.FONT_BLACK_NORMAL).getBold());
            Assertions.assertTrue(fontOf(excel, KanopusExcel.Style.FONT_RED_BOLD).getBold());
            Assertions.assertNotNull(excel.getStyle(KanopusExcel.DefaultStyle.DEFAULT_DATE));
            Assertions.assertNull(excel.getStyle((KanopusExcel.Style) null));
            Assertions.assertNull(excel.getStyle((KanopusExcel.DefaultStyle) null));

            CellStyle custom = excel.createCellStyle();
            Font font = excel.createFont();
            font.setItalic(true);
            custom.setFont(font);
            excel.registerStyle("CUSTOM", custom);
            Assertions.assertSame(custom, excel.getStyle("CUSTOM"));
            Assertions.assertSame(custom, excel.styles.get("CUSTOM"));
        }
    }

    private static Font fontOf(KanopusExcel excel, KanopusExcel.Style style) {
        CellStyle cellStyle = excel.getStyle(style);
        Assertions.assertNotNull(cellStyle);
        return ((XSSFCellStyle) cellStyle).getFont();
    }

    @Test
    void closeIsIdempotentAndWriteReleasesResources() throws Exception {
        KanopusExcel excel = new KanopusExcel();
        excel.createSheet("S").createRow().createCell("x");
        try (OutputStream out = Files.newOutputStream(OUT.resolve("close.xlsx"))) {
            excel.write(out);
        }
        excel.close();
        try (InputStream in = Files.newInputStream(OUT.resolve("close.xlsx")); XSSFWorkbook wb = new XSSFWorkbook(in)) {
            Assertions.assertEquals("x", wb.getSheet("S").getRow(0).getCell(0).getStringCellValue());
        }
        try (KanopusExcel unused = new KanopusExcel()) {
            unused.createSheet("never written");
        }
    }

    @Test
    void autoSizeConsidersFlushedRows() throws Exception {
        ByteArrayOutputStream output;
        try (KanopusExcel excel = new KanopusExcel(10, true)) {
            KSheet sheet = excel.createSheet("RECORDS");

            KRow header = sheet.createRow();
            header.createCell("Text", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
            header.createCell("Integer", KanopusExcel.Style.TABLE_TITLE_INFO);
            header.createCell("Date", KanopusExcel.Style.TABLE_TITLE_INFO);
            header.createCell("LocalDate", KanopusExcel.Style.TABLE_TITLE_INFO);
            header.createCell("LocalDatetime", KanopusExcel.Style.TABLE_TITLE_INFO);
            header.createCell("Boolean", KanopusExcel.Style.TABLE_TITLE_INFO);

            KRow wide = sheet.createRow();
            wide.createCell("A very very very very very long text that should widen the first column");
            for (int i = 0; i < 1000; i++) {
                KRow row = sheet.createRow();
                row.createCell("A" + i);
                row.createCell(i);
                row.createCell(new Date());
                row.createCell(LocalDate.now());
                row.createCell(LocalDateTime.now());
                row.createCell(i % 2 == 0);
            }

            sheet.createFreezePane(0, 1);
            sheet.autoFilter(header.getColumns());
            sheet.autoSize(header.getColumns());
            output = excel.generateOutput();
        }

        try (XSSFWorkbook wb = readBack(output, "products.xlsx")) {
            XSSFSheet sheet = wb.getSheet("RECORDS");
            Assertions.assertEquals(1001, sheet.getLastRowNum());
            Assertions.assertTrue(sheet.getColumnWidth(0) > 30 * 256, "width " + sheet.getColumnWidth(0));
            Assertions.assertTrue(sheet.getCTWorksheet().isSetAutoFilter());
            Assertions.assertEquals("This is the title of the CODE field", sheet.getRow(0).getCell(0).getCellComment().getString().getString());
        }
    }
}
