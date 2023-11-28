package cl.kanopus.excel.writer.iterator;

import cl.kanopus.excel.writer.RecordsExcel;
import cl.kanopus.common.data.Paginator;
import cl.kanopus.common.data.Searcher;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 *
 * @author godheaven
 */
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

        excel.autoSizeColumns();
        excel.autoFilter();

        ByteArrayOutputStream baos = excel.generateOutput();
        ByteArrayInputStream is = new ByteArrayInputStream(baos.toByteArray());
        return is;
    }

    protected abstract Paginator<T> findRecords(Searcher searcher);

    protected abstract List<String> header();

    protected abstract List<Object> row(T object);
}