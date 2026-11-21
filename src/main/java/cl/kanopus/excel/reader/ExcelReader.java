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
package cl.kanopus.excel.reader;

import cl.kanopus.excel.reader.validator.LoadValidatorException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

/**
 * Reads XLS, XLSX and XLSM workbooks, delivering each data row of the sheets that have a registered {@link SheetEventListener} to its {@link RowProcess} as a map {@code header title -> formatted cell value}.
 *
 * <p>
 * The row at {@link SheetEventListener#getStartProcess()} is used as the header row; every following non-empty row is a data row. Blank rows are skipped (they do not stop the processing). Each callback receives a new map instance, so it is safe to keep a reference to it.
 *
 * <p>
 * The workbook format is detected from the content (not the file extension). For large XLSX/XLSM files enable {@link #setStreaming(boolean) streaming}, which parses the sheets with SAX instead of loading the whole workbook in memory (formulas are then read from their cached values instead of being
 * evaluated).
 *
 * <p>
 * For untrusted uploads configure {@link #setMaxInputSize(long)} and, optionally, the JVM-wide POI limits through {@link #configureZipSecurity(double, long)} and {@link #configureByteArrayMaxOverride(int)}.
 *
 * <p>
 * Instances are not thread-safe.
 */
public final class ExcelReader {

    /** Value used to disable the input size limit (default). */
    public static final long UNLIMITED = -1;

    private final ArrayList<SheetEventListener> listeners = new ArrayList<>();

    private int totalProcessedSheets = 0;
    private int totalProcessedRecords = 0;
    private boolean titlesToUpperCase = true;
    private boolean streaming = false;
    private long maxInputSize = UNLIMITED;

    public void addListener(SheetEventListener sheetEventListener) {
        listeners.add(sheetEventListener);
    }

    private SheetEventListener findListener(String sheetName) {
        for (SheetEventListener sheetEventListener : listeners) {
            if (sheetEventListener.getSheetName().equals(sheetName)) {
                return sheetEventListener;
            }
        }
        return null;
    }

    /**
     * Processes every sheet of the given workbook file that has a registered listener. The format (XLS, XLSX, XLSM) is detected from the file content.
     */
    public void processAllSheets(File file) throws ExcelReaderException {
        try {
            checkSize(file.length());
            if (streaming && FileMagic.valueOf(file) == FileMagic.OOXML) {
                try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
                    processStreaming(pkg);
                }
            } else {
                try (Workbook wb = WorkbookFactory.create(file, null, true)) {
                    processWorkbook(wb);
                }
            }
        } catch (ExcelReaderException ex) {
            throw ex;
        } catch (Exception ex) {
            throw wrap(ex);
        }
    }

    /**
     * Processes every sheet of the workbook read from the stream that has a registered listener. The format (XLS, XLSX, XLSM) is detected from the content. The stream is not closed.
     */
    public void processAllSheets(InputStream is) throws ExcelReaderException {
        try {
            InputStream in = FileMagic.prepareToCheckMagic(maxInputSize >= 0 ? new LimitedInputStream(is, maxInputSize) : is);
            if (streaming && FileMagic.valueOf(in) == FileMagic.OOXML) {
                Path tmp = Files.createTempFile("klib-excel-", ".xlsx");
                try {
                    try (OutputStream out = Files.newOutputStream(tmp)) {
                        in.transferTo(out);
                    }
                    try (OPCPackage pkg = OPCPackage.open(tmp.toFile(), PackageAccess.READ)) {
                        processStreaming(pkg);
                    }
                } finally {
                    Files.deleteIfExists(tmp);
                }
            } else {
                try (Workbook wb = WorkbookFactory.create(in)) {
                    processWorkbook(wb);
                }
            }
        } catch (ExcelReaderException ex) {
            throw ex;
        } catch (Exception ex) {
            throw wrap(ex);
        }
    }

    /**
     * @param xlsx
     *            ignored since 4.20.0, the format is detected from the content.
     * @deprecated use {@link #processAllSheets(InputStream)}.
     */
    @Deprecated(since = "4.20.0", forRemoval = true)
    public void processAllSheets(InputStream is, boolean xlsx) throws ExcelReaderException {
        processAllSheets(is);
    }

    private void processWorkbook(Workbook wb) {
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        DataFormatter formatter = new DataFormatter();
        for (int k = 0; k < wb.getNumberOfSheets(); k++) {
            SheetEventListener listener = findListener(wb.getSheetName(k));
            if (listener == null) {
                continue;
            }
            totalProcessedSheets++;
            SheetHandler handler = new SheetHandler(listener);
            Sheet sheet = wb.getSheetAt(k);
            int lastRow = sheet.getLastRowNum();
            for (int r = 0; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                handler.startRow(r);
                for (Iterator<Cell> cells = row.cellIterator(); cells.hasNext();) {
                    Cell cell = cells.next();
                    String value;
                    try {
                        value = formatter.formatCellValue(cell, evaluator);
                    } catch (Exception ex) {
                        throw new ExcelReaderException(ex, r, listener.getSheetName());
                    }
                    handler.cell(cell.getColumnIndex(), value);
                }
                handler.endRow();
            }
        }
    }

    private void processStreaming(OPCPackage pkg) throws Exception {
        XSSFReader xssfReader = new XSSFReader(pkg);
        SharedStrings strings = xssfReader.getSharedStringsTable();
        StylesTable styles = xssfReader.getStylesTable();
        DataFormatter formatter = new DataFormatter();
        XSSFReader.SheetIterator sheets = xssfReader.getSheetIterator();
        while (sheets.hasNext()) {
            try (InputStream sheetData = sheets.next()) {
                SheetEventListener listener = findListener(sheets.getSheetName());
                if (listener == null) {
                    continue;
                }
                totalProcessedSheets++;
                SheetHandler handler = new SheetHandler(listener);
                XMLReader parser = XMLHelper.newXMLReader();
                parser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, handler, formatter, false));
                parser.parse(new InputSource(sheetData));
            }
        }
    }

    private void checkSize(long size) {
        if (maxInputSize >= 0 && size > maxInputSize) {
            throw new ExcelReaderException(sizeExceededMessage());
        }
    }

    private String sizeExceededMessage() {
        return "The excel input exceeds the maximum allowed size of " + maxInputSize + " bytes";
    }

    private ExcelReaderException wrap(Exception ex) {
        for (Throwable t = ex; t != null; t = t.getCause()) {
            if (t instanceof ExcelReaderException ere) {
                return ere;
            }
            if (t instanceof SizeExceededException) {
                return new ExcelReaderException(sizeExceededMessage());
            }
        }
        return new ExcelReaderException(ex);
    }

    /**
     * Configures the JVM-wide zip-bomb protection used by POI when opening OOXML files. Note these settings are global (static) for all POI usages of the JVM.
     *
     * @param minInflateRatio
     *            minimum allowed ratio compressed/uncompressed size of each zip entry (POI default 0.01).
     * @param maxEntrySize
     *            maximum uncompressed size in bytes of each zip entry (POI default 4GB).
     */
    public static void configureZipSecurity(double minInflateRatio, long maxEntrySize) {
        ZipSecureFile.setMinInflateRatio(minInflateRatio);
        ZipSecureFile.setMaxEntrySize(maxEntrySize);
    }

    /**
     * Overrides the JVM-wide maximum size of the byte arrays POI allocates while reading records ({@link IOUtils#setByteArrayMaxOverride(int)}). Use {@code -1} to restore the POI defaults. This setting is global for all POI usages of the JVM.
     */
    public static void configureByteArrayMaxOverride(int maxOverride) {
        IOUtils.setByteArrayMaxOverride(maxOverride);
    }

    public int getTotalProcessedSheets() {
        return totalProcessedSheets;
    }

    public int getTotalProcessedRecords() {
        return totalProcessedRecords;
    }

    public boolean isTitlesToUpperCase() {
        return titlesToUpperCase;
    }

    public void setTitlesToUpperCase(boolean titlesToUpperCase) {
        this.titlesToUpperCase = titlesToUpperCase;
    }

    public boolean isStreaming() {
        return streaming;
    }

    /**
     * Enables the low-memory SAX based reader for XLSX/XLSM inputs (XLS inputs are always read with the in-memory model). In streaming mode formula cells return their cached value instead of being evaluated.
     */
    public void setStreaming(boolean streaming) {
        this.streaming = streaming;
    }

    public long getMaxInputSize() {
        return maxInputSize;
    }

    /**
     * Maximum accepted size in bytes of the (compressed) input file or stream; {@link #UNLIMITED} (default) disables the check. Inputs exceeding the limit fail with an {@link ExcelReaderException}.
     */
    public void setMaxInputSize(long maxInputSize) {
        this.maxInputSize = maxInputSize;
    }

    /** Collects the cells of each row of one sheet and dispatches the data rows to the listener. */
    private final class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        private final SheetEventListener listener;
        private final Map<String, String> titles = new HashMap<>();
        private HashMap<String, String> current;
        private int currentRow;
        private int nextColumn;
        private int countContent;

        SheetHandler(SheetEventListener listener) {
            this.listener = listener;
        }

        @Override
        public void startRow(int rowNum) {
            currentRow = rowNum;
            nextColumn = 0;
            countContent = 0;
            current = new LinkedHashMap<>();
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int column = cellReference != null ? new CellReference(cellReference).getCol() : nextColumn;
            cell(column, formattedValue);
        }

        void cell(int column, String value) {
            nextColumn = column + 1;
            if (currentRow < listener.getStartProcess()) {
                return;
            }
            String key = "col" + column;
            if (currentRow == listener.getStartProcess()) {
                if (value != null) {
                    String title = value.trim();
                    titles.put(key, titlesToUpperCase ? title.toUpperCase(Locale.ROOT) : title);
                }
            } else {
                String title = titles.get(key);
                if (title == null || title.isEmpty()) {
                    // data outside of the header columns
                    return;
                }
                String trimmed = value != null ? value.trim() : null;
                if (trimmed != null && !trimmed.isEmpty()) {
                    countContent++;
                }
                current.put(title, trimmed);
            }
        }

        @Override
        public void endRow(int rowNum) {
            endRow();
        }

        void endRow() {
            if (currentRow <= listener.getStartProcess() || current.isEmpty() || countContent == 0) {
                return;
            }
            totalProcessedRecords++;
            try {
                listener.getRowProcess().process(current);
            } catch (ExcelReaderException ex) {
                throw ex;
            } catch (LoadValidatorException lve) {
                throw new ExcelReaderException(lve, lve.getCode().getMessage(lve.getArgs()), currentRow, listener.getSheetName());
            } catch (Exception ex) {
                throw new ExcelReaderException(ex, currentRow, listener.getSheetName());
            }
        }
    }

    /** Signals that {@link LimitedInputStream} reached the configured limit. */
    private static final class SizeExceededException extends IOException {

        private static final long serialVersionUID = 1L;

        SizeExceededException() {
            super("maximum input size exceeded");
        }
    }

    /** Input stream failing once more than {@code limit} bytes were read. */
    private static final class LimitedInputStream extends java.io.FilterInputStream {

        private final long limit;
        private long count;
        private long mark;

        LimitedInputStream(InputStream in, long limit) {
            super(in);
            this.limit = limit;
        }

        private void add(long n) throws IOException {
            if (n > 0) {
                count += n;
                if (count > limit) {
                    throw new SizeExceededException();
                }
            }
        }

        @Override
        public int read() throws IOException {
            int b = super.read();
            if (b >= 0) {
                add(1);
            }
            return b;
        }

        @Override
        public int read(byte[] b, int off, int len) throws IOException {
            int n = super.read(b, off, len);
            add(n);
            return n;
        }

        @Override
        public long skip(long n) throws IOException {
            long skipped = super.skip(n);
            add(skipped);
            return skipped;
        }

        @Override
        public synchronized void mark(int readlimit) {
            super.mark(readlimit);
            mark = count;
        }

        @Override
        public synchronized void reset() throws IOException {
            super.reset();
            count = mark;
        }

        @Override
        public void close() {
            // the caller owns the wrapped stream
        }
    }
}
