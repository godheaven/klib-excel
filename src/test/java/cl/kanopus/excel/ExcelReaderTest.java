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
import cl.kanopus.excel.reader.ExcelReaderException;
import cl.kanopus.excel.reader.SheetEventListener;
import cl.kanopus.excel.reader.validator.LoadValidatorException;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

class ExcelReaderTest {

    private static final Path OUT = Path.of("target", "excel-reader-test");

    @BeforeAll
    static void setUp() throws Exception {
        Files.createDirectories(OUT);
    }

    private static File resource(String name) throws URISyntaxException {
        return new File(ExcelReaderTest.class.getClassLoader().getResource(name).toURI());
    }

    private static ExcelReader reader(boolean streaming, Map<String, List<HashMap<String, String>>> collected, String... sheets) {
        ExcelReader reader = new ExcelReader();
        reader.setStreaming(streaming);
        for (String sheet : sheets) {
            List<HashMap<String, String>> rows = new ArrayList<>();
            collected.put(sheet, rows);
            reader.addListener(new SheetEventListener(sheet, rows::add));
        }
        return reader;
    }

    /**
     * DATA: header + 2 rows, 3 blank rows (one physically present with empty cells), a last row and a value outside of the header. OTHER: a different header. IGNORED: no listener.
     */
    private static void fillWorkbook(Workbook wb) {
        Sheet data = wb.createSheet("DATA");
        row(data, 0, "Code", "Name", " Qty ");
        row(data, 1, "A1", "Alpha", "1");
        row(data, 2, "A2", "Beta", "2");
        row(data, 4, "", "  ", "");
        Row last = row(data, 7, "A3", "Ñandú", "3");
        last.createCell(5).setCellValue("outside header");
        data.createRow(8).createCell(0).setCellFormula("2*21");

        Sheet other = wb.createSheet("OTHER");
        row(other, 0, "Id");
        row(other, 1, "X");

        row(wb.createSheet("IGNORED"), 0, "Nothing");
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
    }

    private static Row row(Sheet sheet, int index, String... values) {
        Row row = sheet.createRow(index);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
        return row;
    }

    private static File write(Workbook wb, String name) throws Exception {
        try (wb) {
            fillWorkbook(wb);
            File file = OUT.resolve(name).toFile();
            try (OutputStream out = Files.newOutputStream(file.toPath())) {
                wb.write(out);
            }
            return file;
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"test_reader.xls", "test_reader.xlsx"})
    void readResources(String name) throws Exception {
        for (boolean streaming : new boolean[]{false, true}) {
            Map<String, List<HashMap<String, String>>> collected = new HashMap<>();
            ExcelReader reader = reader(streaming, collected, "RECORDS");
            reader.processAllSheets(resource(name));

            Assertions.assertEquals(1, reader.getTotalProcessedSheets());
            Assertions.assertEquals(9, reader.getTotalProcessedRecords());
            List<HashMap<String, String>> rows = collected.get("RECORDS");
            Assertions.assertEquals(9, rows.size());
            Assertions.assertEquals("av2231fwg", rows.get(0).get("CODE"));
            Assertions.assertEquals("Brown Purse", rows.get(0).get("NAME"));
            Assertions.assertEquals("0", rows.get(0).get("QUANTITY"));
            Assertions.assertEquals("v435nn85n", rows.get(7).get("CODE"));
            Assertions.assertEquals("42", rows.get(7).get("QUANTITY"));
            // every row gets its own map
            Assertions.assertNotSame(rows.get(0), rows.get(1));
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"xls", "xlsx", "xlsm"})
    void readsAllRowsAfterBlankRowsAndResetsTitlesPerSheet(String format) throws Exception {
        Workbook wb = switch (format) {
            case "xls" -> new HSSFWorkbook();
            case "xlsm" -> new XSSFWorkbook(XSSFWorkbookType.XLSM);
            default -> new XSSFWorkbook();
        };
        File file = write(wb, "blank-rows." + format);

        for (boolean streaming : new boolean[]{false, true}) {
            for (boolean fromStream : new boolean[]{false, true}) {
                Map<String, List<HashMap<String, String>>> collected = new HashMap<>();
                ExcelReader reader = reader(streaming, collected, "DATA", "OTHER");
                if (fromStream) {
                    try (InputStream is = Files.newInputStream(file.toPath())) {
                        reader.processAllSheets(is);
                    }
                } else {
                    reader.processAllSheets(file);
                }
                String mode = format + " streaming=" + streaming + " stream=" + fromStream;
                Assertions.assertEquals(2, reader.getTotalProcessedSheets(), mode);
                Assertions.assertEquals(5, reader.getTotalProcessedRecords(), mode);

                List<HashMap<String, String>> data = collected.get("DATA");
                Assertions.assertEquals(4, data.size(), mode);
                Assertions.assertEquals(Map.of("CODE", "A1", "NAME", "Alpha", "QTY", "1"), data.get(0), mode);
                Assertions.assertEquals(Map.of("CODE", "A2", "NAME", "Beta", "QTY", "2"), data.get(1), mode);
                Assertions.assertEquals(Map.of("CODE", "A3", "NAME", "Ñandú", "QTY", "3"), data.get(2), mode);
                Assertions.assertEquals(Map.of("CODE", "42"), data.get(3), mode);
                Assertions.assertFalse(data.get(2).containsKey(null), mode);

                Assertions.assertEquals(List.of(Map.of("ID", "X")), collected.get("OTHER"), mode);
            }
        }
    }

    @Test
    void startProcessAndTitlesCase() throws Exception {
        File file = OUT.resolve("start-process.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file.toPath())) {
            Sheet sheet = wb.createSheet("S");
            row(sheet, 0, "Report title");
            row(sheet, 2, "Code", "Name");
            row(sheet, 3, "1", "One");
            wb.write(out);
        }
        for (boolean streaming : new boolean[]{false, true}) {
            List<HashMap<String, String>> rows = new ArrayList<>();
            ExcelReader reader = new ExcelReader();
            reader.setStreaming(streaming);
            reader.setTitlesToUpperCase(false);
            Assertions.assertFalse(reader.isTitlesToUpperCase());
            Assertions.assertEquals(streaming, reader.isStreaming());
            reader.addListener(new SheetEventListener("S", rows::add, 2));
            reader.processAllSheets(file);
            Assertions.assertEquals(List.of(Map.of("Code", "1", "Name", "One")), rows);
        }
    }

    @Test
    @SuppressWarnings("removal")
    void deprecatedFlagIsIgnoredAndFormatIsDetected() throws Exception {
        List<HashMap<String, String>> rows = new ArrayList<>();
        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("RECORDS", rows::add));
        try (InputStream is = Files.newInputStream(resource("test_reader.xlsx").toPath())) {
            reader.processAllSheets(is, false);
        }
        Assertions.assertEquals(9, rows.size());
    }

    @Test
    void validatorErrorsAreReportedWithRowAndSheet() throws Exception {
        for (boolean streaming : new boolean[]{false, true}) {
            ExcelReader reader = new ExcelReader();
            reader.setStreaming(streaming);
            reader.addListener(new SheetEventListener("RECORDS", row -> {
                if ("mbvjkgip5".equals(row.get("CODE"))) {
                    throw new LoadValidatorException(LoadValidatorException.ErrorCode.VALUE_REQUIRED, "CODE");
                }
            }));
            ExcelReaderException ex = Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(resource("test_reader.xlsx")));
            Assertions.assertEquals("El valor del campo [CODE] es requerido", ex.getMessage());
            Assertions.assertEquals("RECORDS", ex.getSheetName());
            Assertions.assertEquals(4, ex.getCurrentRow());
            Assertions.assertInstanceOf(LoadValidatorException.class, ex.getCause());
        }
    }

    @Test
    void processErrorsAreWrapped() throws Exception {
        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("RECORDS", row -> {
            throw new IllegalStateException("boom");
        }));
        ExcelReaderException ex = Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(resource("test_reader.xls")));
        Assertions.assertInstanceOf(IllegalStateException.class, ex.getCause());
        Assertions.assertEquals(2, ex.getCurrentRow());
    }

    @Test
    void invalidContentFails() throws Exception {
        File file = OUT.resolve("not-excel.xlsx").toFile();
        Files.writeString(file.toPath(), "this is not a workbook");
        for (boolean streaming : new boolean[]{false, true}) {
            ExcelReader reader = new ExcelReader();
            reader.setStreaming(streaming);
            Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(file));
            Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(new ByteArrayInputStream("nope".getBytes())));
        }
        ExcelReader reader = new ExcelReader();
        Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(OUT.resolve("missing.xlsx").toFile()));
    }

    @Test
    void maxInputSize() throws Exception {
        File file = resource("test_reader.xlsx");
        for (boolean streaming : new boolean[]{false, true}) {
            ExcelReader reader = new ExcelReader();
            reader.setStreaming(streaming);
            Assertions.assertEquals(ExcelReader.UNLIMITED, reader.getMaxInputSize());
            reader.setMaxInputSize(1024);
            Assertions.assertEquals(1024, reader.getMaxInputSize());

            ExcelReaderException ex = Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(file));
            Assertions.assertTrue(ex.getMessage().contains("1024 bytes"), ex.getMessage());
            try (InputStream is = Files.newInputStream(file.toPath())) {
                ex = Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(is));
                Assertions.assertTrue(ex.getMessage().contains("maximum allowed size"), ex.getMessage());
            }
            try (InputStream is = Files.newInputStream(resource("test_reader.xls").toPath())) {
                Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(is));
            }

            List<HashMap<String, String>> rows = new ArrayList<>();
            reader.addListener(new SheetEventListener("RECORDS", rows::add));
            reader.setMaxInputSize(file.length());
            reader.processAllSheets(file);
            try (InputStream is = Files.newInputStream(file.toPath())) {
                reader.processAllSheets(is);
            }
            Assertions.assertEquals(18, rows.size());
        }
    }

    @Test
    void zipSecuritySettings() throws Exception {
        File file = OUT.resolve("compressible.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file.toPath())) {
            Sheet sheet = wb.createSheet("S");
            row(sheet, 0, "Text");
            for (int i = 1; i < 2000; i++) {
                row(sheet, i, "same value repeated");
            }
            wb.write(out);
        }
        double ratio = ZipSecureFile.getMinInflateRatio();
        long entry = ZipSecureFile.getMaxEntrySize();
        try {
            ExcelReader.configureZipSecurity(0.5, entry);
            ExcelReader reader = new ExcelReader();
            reader.addListener(new SheetEventListener("S", row -> {
            }));
            Assertions.assertThrows(ExcelReaderException.class, () -> reader.processAllSheets(file));
            ExcelReader.configureByteArrayMaxOverride(1_000_000);
        } finally {
            ExcelReader.configureZipSecurity(ratio, entry);
            ExcelReader.configureByteArrayMaxOverride(-1);
        }
        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("S", row -> {
        }));
        reader.processAllSheets(file);
        Assertions.assertEquals(1999, reader.getTotalProcessedRecords());
    }
}
