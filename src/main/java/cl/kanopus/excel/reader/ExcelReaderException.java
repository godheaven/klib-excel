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
