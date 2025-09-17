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

import cl.kanopus.excel.writer.KanopusExcel;
import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;

import java.util.List;

class RecordsExcel extends KanopusExcel {

    private static final String SHEET_RECORDS = "RECORDS";

    private int columns = 0;
    private int totalRecords = 0;
    private boolean enableRecordStyle = false;
    private final KSheet sheet;

    public RecordsExcel() {
        super();
        this.sheet = super.createSheet(SHEET_RECORDS);
    }

    public RecordsExcel(int rowAccessWindowSize, boolean compressTmpFiles) {
        super(rowAccessWindowSize, compressTmpFiles);
        this.sheet = super.createSheet(SHEET_RECORDS);
    }

    public void createHeaders(List<String> titles) {
        KRow header = sheet.createRow();
        for (String title : titles) {
            header.createCell(title.toUpperCase(), Style.TABLE_TITLE_NORMAL);
        }
        sheet.createFreezePane(0, 1);
        columns = titles.size();
    }

    public void createRecord(List<Object> values) {
        KRow row = sheet.createRow();
        for (Object value : values) {
            row.createCell(value, enableRecordStyle ? Style.TABLE_VALUE_NORMAL : null);
        }
        this.totalRecords++;
    }

    public void autoSize() {
        sheet.autoFilter(columns);
    }

    public void autoFilter() {
        sheet.autoSize(columns);
    }

    public boolean isEnableRecordStyle() {
        return enableRecordStyle;
    }

    public void setEnableRecordStyle(boolean enableRecordStyle) {
        this.enableRecordStyle = enableRecordStyle;
    }

    public int getTotalRecords() {
        return totalRecords;
    }

}
