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

public class SheetEventListener {

    private final String sheetName;
    private final RowProcess rowProcess;
    private final int startProcess;

    public SheetEventListener(String sheetName, RowProcess rowProcess) {
        this.sheetName = sheetName;
        this.rowProcess = rowProcess;
        this.startProcess = 0;
    }

    public SheetEventListener(String sheetName, RowProcess rowProcess, int startProcess) {
        this.sheetName = sheetName;
        this.rowProcess = rowProcess;
        this.startProcess = startProcess;
    }

    public String getSheetName() {
        return sheetName;
    }

    public RowProcess getRowProcess() {
        return rowProcess;
    }

    public int getStartProcess() {
        return startProcess;
    }

}
