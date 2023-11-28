package cl.kanopus.excel.reader;

import java.util.HashMap;

public interface RowProcess {

    void process(HashMap<String, String> row) throws Exception;

}
