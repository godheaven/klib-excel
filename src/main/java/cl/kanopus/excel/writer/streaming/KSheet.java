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
package cl.kanopus.excel.writer.streaming;

import cl.kanopus.excel.writer.KanopusExcel;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class KSheet {

    private static final int AUTO_LIMIT = 50000;
    private int indexRow = 0;
    private final SXSSFSheet sheet;
    private final KanopusExcel excel;

    public KSheet(KanopusExcel excel, SXSSFSheet sheet) {
        this.excel = excel;
        this.sheet = sheet;
        this.sheet.setDefaultColumnWidth(15);
    }

    public KRow createRow() {
        return new KRow(excel, sheet.createRow(indexRow++));
    }

    public KRow createRow(int rownum) {
        indexRow = rownum;
        return createRow();
    }

    public void createFreezePane(int colSplit, int rowSplit) {
        sheet.createFreezePane(colSplit, rowSplit);
    }

    public void autoFilter(int columns) {
        autoFilter(0, columns);
    }
    
    public void autoFilter(int firstRow, int columns) {
        if (columns > 0 && indexRow > 0 && indexRow < AUTO_LIMIT) {
            sheet.setAutoFilter(new CellRangeAddress(firstRow, indexRow - 1, 0, columns - 1));
        }
    }

    public void autoSize(int columns) {

        if (columns > 0 && indexRow < AUTO_LIMIT) {
            for (int i = 0; i < columns; i++) {
                sheet.trackAllColumnsForAutoSizing();
                sheet.autoSizeColumn(i, false);
            }
        }

    }

}
