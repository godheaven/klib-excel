package cl.kanopus.excel.writer;

import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class ProductsExcel extends AbstractExcel {

    public static final String SHEET_PRODUCTS = "PRODUCTS";
    
    public enum Field {

        CODE,
        NAME,
        PRICE;

    }

    public void createSheetProducts(Iterator<ProductTO> products) {

        int indexRow = 0;
        HSSFSheet sheet = wb.createSheet(SHEET_PRODUCTS);

        this.row = sheet.createRow(indexRow++);
        int column = 0;
        createCell(column++, Field.CODE.name(), Style.TABLE_TITLE_REQUIRED, "This is the title of the CODE field");
        createCell(column++, Field.NAME.name(), Style.TABLE_TITLE_REQUIRED, "This is the title of the NAME field");
        createCell(column++, Field.PRICE.name(), Style.TABLE_TITLE_OPTIONAL, "This is the title of the PRICE field");

        while (products.hasNext()) {
            ProductTO p = products.next();
            this.row = sheet.createRow(indexRow++);
            column = 0;
            createCell(column++, p.getCode(), Style.TABLE_VALUE_INSERT);
            createCell(column++, p.getName(), Style.TABLE_VALUE_INSERT);
            createCell(column++, p.getPrice(), Style.TABLE_VALUE_INSERT);
        }

        for (int i = 0; i < column; i++) {
            sheet.autoSizeColumn(i);
        }
        sheet.createFreezePane(0, 1);

    }

}
