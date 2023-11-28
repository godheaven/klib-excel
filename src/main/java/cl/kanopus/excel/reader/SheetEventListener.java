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
