package cl.kanopus.excel.writer;

import java.util.List;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class RecordsExcel extends AbstractExcel {

    public static final String SHEET_REGISTROS = "REGISTROS";

    private final HSSFSheet sheet;
    private int indexRow = 0;
    private int columns = 0;

    public RecordsExcel() {
        this.sheet = wb.createSheet(SHEET_REGISTROS);
    }

    public void createHeaders(List<String> titles) {
        this.row = sheet.createRow(indexRow++);
        this.columns = titles.size();
        int c = 0;
        for (String title : titles) {
            createCell(c++, title.toUpperCase(), Style.TABLE_TITLE_NORMAL);
        }
        sheet.createFreezePane(0, 1);
    }

    public void createRecord(List<Object> values) {
        this.row = sheet.createRow(indexRow++);
        int c = 0;
        for (Object value : values) {
            createCell(c++, value, Style.TABLE_VALUE_NORMAL);
        }
    }

    public void autoSizeColumns() {
        for (int i = 0; i < columns; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    public void autoFilter() {
        if (indexRow > 0) {
            sheet.setAutoFilter(new CellRangeAddress(0, indexRow - 1, 0, columns - 1));
        }
    }
}
