package cl.kanopus.excel.reader;

import cl.kanopus.excel.reader.validator.LoadValidatorException;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * File for HSSF testing/examples
 *
 * THIS IS NOT THE MAIN HSSF FILE!! This is a utility for testing functionality.
 * It does contain sample API usage that may be educational to regular API
 * users.
 */
public final class ExcelReader {

    private final ArrayList<SheetEventListener> listeners = new ArrayList<>();

    private int totalProcessedSheets = 0;
    private int totalProcessedRecords = 0;
    private boolean titlesToUpperCase = true;
    private final HashMap<String, String> hashTitles = new HashMap<>();
    private final HashMap<String, String> hashRow = new HashMap<>();

    public void addListener(SheetEventListener sheetEventListener) {
        listeners.add(sheetEventListener);
    }

    private SheetEventListener findListener(String sheetName) {
        SheetEventListener listener = null;
        for (SheetEventListener sheetEventListener : listeners) {
            if (sheetEventListener.getSheetName().equals(sheetName)) {
                listener = sheetEventListener;
                break;
            }
        }
        return listener;
    }

    public void processAllSheets(File file) throws ExcelReaderException {
        try {
            try (FileInputStream fis = new FileInputStream(file);) {
                processAllSheets(fis);
            }
        } catch (ExcelReaderException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new ExcelReaderException(ex);
        }
    }

    public void processAllSheets(InputStream is) throws ExcelReaderException {

        try {

            HSSFWorkbook wb = null;
            try {
                wb = new HSSFWorkbook(is);

                for (int k = 0; k < wb.getNumberOfSheets(); k++) {

                    SheetEventListener listener = findListener(wb.getSheetName(k));
                    if (listener != null) {
                        totalProcessedSheets++;
                        HSSFSheet sheet = wb.getSheetAt(k);
                        int rows = sheet.getPhysicalNumberOfRows();

                        for (int r = 0; r < rows; r++) {
                            HSSFRow row = sheet.getRow(r);
                            if (row == null || r < listener.getStartProcess()) {
                                continue;
                            }

                            int countContent = 0;
                            DataFormatter formatter = new DataFormatter();
                            for (Iterator<Cell> cells = row.cellIterator(); cells.hasNext();) {
                                try {
                                    Cell cell = cells.next();

                                    String key = "col" + cell.getColumnIndex();
                                    String value = formatter.formatCellValue(cell);

                                    if (r == listener.getStartProcess()) {
                                        if (value != null) {
                                            hashTitles.put(key, titlesToUpperCase ? value.toUpperCase().trim() : value.trim());
                                        }
                                    } else {
                                        countContent += (value != null && !value.trim().isEmpty()) ? 1 : 0;
                                        hashRow.put(hashTitles.get(key), value != null ? value.trim() : null);
                                    }
                                } catch (Exception ex) {
                                    throw new ExcelReaderException(ex, row.getRowNum(), listener.getSheetName());
                                }
                            }

                            if (!hashRow.isEmpty() && countContent > 0) {
                                totalProcessedRecords++;
                                try {
                                    listener.getRowProcess().process(hashRow);
                                } catch (LoadValidatorException lve) {
                                    throw new ExcelReaderException(lve, lve.getCode().getMessage(lve.getArgs()), row.getRowNum(), listener.getSheetName());
                                } catch (Exception ex) {
                                    throw new ExcelReaderException(ex, row.getRowNum(), listener.getSheetName());
                                }
                                hashRow.clear();
                            }
                        }
                    }

                }

            } finally {
                if (wb != null) {
                    wb.close();
                }
            }
        } catch (ExcelReaderException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new ExcelReaderException(ex);
        }
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

}
