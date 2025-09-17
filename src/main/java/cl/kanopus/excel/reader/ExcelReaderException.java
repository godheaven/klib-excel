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

public class ExcelReaderException extends RuntimeException {

	private static final long serialVersionUID = 1L;
	private final int currentRow;
    private final String sheetName;

    public ExcelReaderException(Exception ex) {
        super(ex);
        this.currentRow = 0;
        this.sheetName = "";
    }

    public ExcelReaderException(Exception ex, int currentRow, String sheetName) {
        super(ex);
        this.currentRow = currentRow + 1;
        this.sheetName = sheetName;
    }

    public ExcelReaderException(Exception ex, String message, int currentRow, String sheetName) {
        super(message, ex);
        this.currentRow = currentRow + 1;
        this.sheetName = sheetName;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public String getSheetName() {
        return sheetName;
    }

}
