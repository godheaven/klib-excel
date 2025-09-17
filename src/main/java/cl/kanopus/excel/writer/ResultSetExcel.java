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
package cl.kanopus.excel.writer;

import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;

import java.util.Iterator;
import java.util.Map;

public class ResultSetExcel extends KanopusExcel {

    private static final String SHEET_RECORDS = "RECORDS";
    private boolean enableAutoFilter = true;
    private boolean enableFreezePane = true;

    public ResultSetExcel() {
    }

    public ResultSetExcel(int rowAccessWindowSize, boolean compressTmpFiles) {
        super(rowAccessWindowSize, compressTmpFiles);
    }

    public int createSheet(Iterator<Map<String, Object>> records) {
        return createSheet(SHEET_RECORDS, records);
    }

    public int createSheet(String sheetName, Iterator<Map<String, Object>> records) {
        int rows = 0;
        KSheet sheet = createSheet(SHEET_RECORDS);

        boolean headerIncluded = false;

        int columns = 0;
        while (records.hasNext()) {
            Map<String, Object> p = records.next();

            if (!headerIncluded) {
                KRow header = sheet.createRow();
                for (String title : p.keySet()) {
                    header.createCell(title, KanopusExcel.Style.TABLE_TITLE_NORMAL);
                    columns++;
                }
                headerIncluded = true;
            }

            KRow row = sheet.createRow();
            for (Object value : p.values()) {
                row.createCell(value, KanopusExcel.Style.TABLE_VALUE_NORMAL);
            }

        }

        if (enableAutoFilter) {
            sheet.autoFilter(columns);
        }

        if (enableFreezePane) {
            sheet.createFreezePane(0, 1);
        }
        return rows;
    }

    public void setEnableAutoFilter(boolean enableAutoFilter) {
        this.enableAutoFilter = enableAutoFilter;
    }

    public void setEnableFreezePane(boolean enableFreezePane) {
        this.enableFreezePane = enableFreezePane;
    }

}
