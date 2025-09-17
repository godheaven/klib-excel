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

import cl.kanopus.common.data.Paginator;
import cl.kanopus.common.data.Searcher;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public abstract class ExcelReaderIterator<T> {

    public InputStream read(Searcher searcher) throws IOException {
        RecordsExcel excel = new RecordsExcel();
        excel.createHeaders(header());

        int offset = 0;
        searcher.setLimit(1000);
        searcher.setOffset(0);
        Paginator<T> paginator = findRecords(searcher);
        while (!paginator.getRecords().isEmpty()) {
            for (T f : paginator.getRecords()) {
                offset++;
                excel.createRecord(row(f));
            }

            if (offset > 0) {
                searcher.setOffset(offset);
                paginator = findRecords(searcher);
            }
        }

        excel.autoSize();
        excel.autoFilter();

        ByteArrayOutputStream baos = excel.generateOutput();
        return new ByteArrayInputStream(baos.toByteArray());
    }

    protected abstract Paginator<T> findRecords(Searcher searcher);

    protected abstract List<String> header();

    protected abstract List<Object> row(T object);
}
