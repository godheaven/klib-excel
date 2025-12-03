![Logo](https://www.kanopus.cl/admin/javax.faces.resource/images/logo-gray.png.xhtml?ln=paradise-layout)

# klib-excel

This project is designed as a utility that allows you to generate Excel files in an easier way.
This simplifies the way you interact with the POI library and reduces the amount of complex code for Excel generation.

## Features

- Allows you to read excel files in a simplified way
- Allows you to write excel files in a simplified way

## 🚀 Installation

Add the dependency to your `pom.xml`:

```xml

<dependency>
	<groupId>cl.kanopus.util</groupId>
	<artifactId>klib-excel</artifactId>
	<version>3.58.0</version>
</dependency>
```

---

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
package cl.kanopus.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;

import org.junit.jupiter.api.Test;
import cl.kanopus.common.util.DesktopUtils;
import cl.kanopus.common.util.FileUtils;
import cl.kanopus.excel.writer.KRow;
import cl.kanopus.excel.writer.KSheet;
import cl.kanopus.excel.writer.KanopusExcel;

public class ExcelWriterTest {

    @Test
    public void generateExcelOneMillionRecords() throws Exception {
        //1millon --> 11seconds --> 25 MB --> 5.000 in memory (default contructor)

        KanopusExcel excel = new KanopusExcel();
        KSheet sheet = excel.createSheet("RECORDS");

        // HEADER
        KRow header = sheet.createRow();
        header.createCell("Text", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Integer", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Date", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("LocalDate", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("LocalDatetime", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
        header.createCell("Boolean", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");

        // RECORDS
        for (int i = 0; i < 1000; i++) {
            KRow row = sheet.createRow();
            row.createCell("A123456789-" + i);
            row.createCell(i);
            row.createCell(new Date());
            row.createCell(LocalDate.now());
            row.createCell(LocalDateTime.now());
            row.createCell(i % 2 == 0);
        }

        sheet.createFreezePane(0, 1);
        sheet.autoFilter(header.getColumns());
        sheet.autoSize(header.getColumns());

        ByteArrayOutputStream baos = excel.generateOutput();
        File file = FileUtils.createFile(baos, "products.xlsx");

        DesktopUtils.open(file);

    }

}
```

## Authors

- [@pabloandres.diazsaavedra](https://www.linkedin.com/in/pablo-diaz-saavedra-4b7b0522/)

## License

This software is licensed under the Apache License, Version 2.0. See the LICENSE file for details.
I hope you enjoy it.

[![Apache License, Version 2.0](https://img.shields.io/badge/license-Apache%20License%202.0-blue.svg)](https://opensource.org/license/apache-2-0)

## Support

For support, email soporte@kanopus.cl
