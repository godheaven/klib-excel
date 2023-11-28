
![Logo](https://www.kanopus.cl/admin/javax.faces.resource/images/logo-gray.png.xhtml?ln=paradise-layout)


# klib-excel

This project is designed as a utility that allows you to generate Excel files in an easier way.
This simplifies the way you interact with the POI library and reduces the amount of complex code for Excel generation.

## Features
- Allows you to read excel files in a simplified way
- Allows you to write excel files in a simplified way 

## Usage/Examples

1. Reading Excel
```java
package cl.kanopus.excel;

import cl.kanopus.excel.reader.ExcelReader;
import cl.kanopus.excel.reader.RowProcess;
import cl.kanopus.excel.reader.SheetEventListener;
import java.io.File;
import java.net.URL;
import java.util.HashMap;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class ExcelReaderTest {

    @Test
    public void readExcel() throws Exception {

        ExcelReader reader = new ExcelReader();
        reader.addListener(new SheetEventListener("RECORDS", new RowProcess() {
            @Override
            public void process(HashMap<String, String> row) throws Exception {
                System.out.println("row-->" + row);
            }
        }));

        URL resource = ExcelReaderTest.class.getClassLoader().getResource("test_reader.xls");
        reader.processAllSheets(new File(resource.getFile()));

        Assertions.assertEquals(1, reader.getTotalProcessedSheets());
        Assertions.assertEquals(9, reader.getTotalProcessedRecords());
    }

}

```

2. Writing Excel
```java
package cl.kanopus.excel.writer;

public class ProductTO {

    private final String code;
    private final String name;
    private final double price;

    public ProductTO(String code, String name, double price) {
        this.code = code;
        this.name = name;
        this.price = price;
    }

    //Getter

}
```

```java
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

```

```java
package cl.kanopus.excel;

import cl.kanopus.common.util.DesktopUtils;
import cl.kanopus.common.util.FileUtils;
import cl.kanopus.excel.writer.ProductTO;
import cl.kanopus.excel.writer.ProductsExcel;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import org.junit.jupiter.api.Test;

public class ExcelWriterTest {

    @Test
    public void testGenerateExcel() throws Exception {

        ArrayList<ProductTO> products = new ArrayList<>();
        products.add(new ProductTO("A123456789", "ASPIRINA", 1000));
        products.add(new ProductTO("B123456789", "BETAMETASONA", 1000));
        products.add(new ProductTO("C123456789", "CLORFENAMINA", 1000));
        products.add(new ProductTO("D123456789", "DERMOCREAM", 1000));
        products.add(new ProductTO("E123456789", "EUCALIPTUS", 1000));
        
        ProductsExcel excel = new ProductsExcel();
        excel.createSheetProducts(products.iterator());

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products.xlsx");
 
        DesktopUtils.open(file);

    }

}
```

## Authors

- [@pabloandres.diazsaavedra](https://www.linkedin.com/in/pablo-diaz-saavedra-4b7b0522/)


## License

This is free software and I hope you enjoy it.

[![GPLv3 License](https://img.shields.io/badge/License-GPL%20v3-yellow.svg)](https://opensource.org/licenses/)





## Support

For support, email pabloandres.diazsaavedra@gmail.com

