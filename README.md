<p style="text-align:left">
  <img src="https://www.kanopus.cl/assets/kanopus_black.png" width="220" alt="Kanopus logo"/>
</p>

![Maven](https://img.shields.io/maven-central/v/cl.kanopus.util/klib-excel) ![License](https://img.shields.io/badge/license-Apache%20License%202.0-blue) ![Java](https://img.shields.io/badge/java-21+-orange)

# klib-excel

This project is designed as a utility that allows you to generate Excel files in an easier way.
This simplifies the way you interact with the POI library and reduces the amount of complex code for Excel generation.

## ✨ Features

- Allows you to read excel files in a simplified way
- Allows you to write excel files in a simplified way

## 🚀 Installation

Add the dependency to your `pom.xml`:

```xml

<dependency>
	<groupId>cl.kanopus.util</groupId>
	<artifactId>klib-excel</artifactId>
	<version>4.20.0</version>
</dependency>
```

---

## 🚀 Usage Guide

1. Reading Excel (XLS, XLSX and XLSM, detected from the content)

```java
import cl.kanopus.excel.reader.ExcelReader;
import cl.kanopus.excel.reader.SheetEventListener;
import java.io.File;

ExcelReader reader = new ExcelReader();
reader.setMaxInputSize(20 * 1024 * 1024); // optional: reject inputs over 20 MB (untrusted uploads)
reader.setStreaming(true);                // optional: low-memory SAX reader for large XLSX/XLSM files
reader.addListener(new SheetEventListener("RECORDS", row -> {
    // row: header title (upper case by default) -> trimmed formatted value; a new map per row
    String code = row.get("CODE");
}));
reader.processAllSheets(new File("products.xlsx"));
```

The row at `startProcess` (default `0`) is the header; blank rows are skipped. Values are formatted with POI's
`DataFormatter` (locale dependent). In streaming mode formula cells return their cached value.

For untrusted files you can also tighten the **JVM-wide** POI limits:

```java
ExcelReader.configureZipSecurity(0.05, 100L * 1024 * 1024); // min inflate ratio, max uncompressed entry size
ExcelReader.configureByteArrayMaxOverride(50_000_000);      // IOUtils.setByteArrayMaxOverride
```

2. Writing Excel

```java
import cl.kanopus.excel.writer.KanopusExcel;
import cl.kanopus.excel.writer.streaming.KRow;
import cl.kanopus.excel.writer.streaming.KSheet;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

try (KanopusExcel excel = new KanopusExcel();
     OutputStream out = Files.newOutputStream(Path.of("products.xlsx"))) {
    KSheet sheet = excel.createSheet("RECORDS");

    KRow header = sheet.createRow();
    header.createCell("Text", KanopusExcel.Style.TABLE_TITLE_INFO, "This is the title of the CODE field");
    header.createCell("Amount", KanopusExcel.Style.TABLE_TITLE_INFO);
    header.createCell("Date", KanopusExcel.Style.TABLE_TITLE_INFO);

    for (int i = 0; i < 1000; i++) {
        KRow row = sheet.createRow();
        row.createCell("A123456789-" + i);
        row.createCell(new java.math.BigDecimal("1234.50")); // numeric cell
        row.createCell(LocalDate.now());                      // date cell
    }

    sheet.createFreezePane(0, 1);
    sheet.autoFilter(header.getColumns());
    sheet.autoSize(header.getColumns());

    excel.write(out); // or excel.generateOutput() for a ByteArrayOutputStream
}
```

3. Writing a list of maps (`ResultSetExcel`)

```java
ResultSetExcel excel = new ResultSetExcel();
int rows = excel.createSheet("Customers", records.iterator()); // header/column order from the first record's keys
```

## ⚠️ Behavior changes in 4.20.0

**Reader (`ExcelReader`)**
- Rows after blank rows are no longer lost (iterates up to `getLastRowNum()`).
- Each `RowProcess` call receives a **new** map (previously the same map instance was cleared and reused).
- Header titles are reset per sheet; cells outside the header columns are ignored (no more `null` key).
- Format detection uses the content (`WorkbookFactory`), so `.xlsm` and misnamed files work;
  `processAllSheets(InputStream, boolean)` is deprecated (the flag is ignored), use `processAllSheets(InputStream)`.
- Titles are upper-cased with `Locale.ROOT`.
- New: `setStreaming(boolean)`, `setMaxInputSize(long)`, `configureZipSecurity(...)`, `configureByteArrayMaxOverride(int)`.

**Validator (`LoadValidator`)**
- `validateRequired` accepts `"0"` and rejects blank (whitespace-only) values.
- Empty optional values: `parseDate`, `parseLong`, `parseInteger`, `parseMoneyToLong`, `parseBoolean`, `parseRut`,
  `parseGTIN` return `null`; `parseString*` still return `""`. Values are trimmed before parsing/validation.
- `parseMoneyToLong` uses the Chilean format: `.` thousands separator, `,` decimal separator (`"$ 1.234.567"`,
  `"1.234,00"`). A single `.` not followed by groups of three digits is a decimal point (`"1000.3"`, as produced by
  Excel), and when both separators are present the last one is the decimal separator. Amounts with non-zero
  decimals are **rejected** (`VALUE_NUMBER_FORMAT_EXCEPTION`) unless `setMoneyRoundingMode(RoundingMode)` is set.
  Previously `"1000.3"` silently became `10003`. Note a single comma is decimal: `"1,000"` is `1`.

**Writer**
- `KanopusExcel` implements `AutoCloseable`; `write(OutputStream)` added. `write`/`generateOutput` release the
  temporary files even on failure (a workbook is single-use).
- The public `styles` field is deprecated: use `getStyle(...)`, `registerStyle(...)`, `createCellStyle()`, `createFont()`.
- `Style.FONT_BLACK_NORMAL` is no longer bold.
- `KRow`: `BigDecimal`, `BigInteger`, `Float`, `Short` (any `Number`) are written as numeric cells instead of text;
  `Calendar` gets the date style and `java.sql.Timestamp` the date-time style.
- `KSheet.autoSize` now considers every row (columns are tracked from the first row; tracking stops after 50,000 rows
  where auto-size is not supported).
- `ResultSetExcel.createSheet(sheetName, ...)` uses the given sheet name (it always used `RECORDS`), returns the number
  of data rows (it always returned `0`) and writes values in the header's column order (missing keys -> empty cell).
- `ExcelReaderIterator`: `autoSize`/`autoFilter` were swapped (no column auto-size was applied); new
  `write(searcher, OutputStream)`; `read` no longer copies the generated bytes twice.

## 👤 Author

⭐**Pablo Andrés Díaz Saavedra** — Founder of **Kanopus – Software Guided by the Stars**⭐

Kanopus is building a constellation of developers creating tools, libraries and platforms that simplify software
engineering.

[GitHub](https://github.com/godheaven) | [LinkedIn](https://www.linkedin.com/in/pablo-diaz-saavedra-4b7b0522/) | [Website](https://kanopus.cl)

## 📄 License

This software is licensed under the Apache License, Version 2.0. See the LICENSE file for details.
I hope you enjoy it.

[![Apache License, Version 2.0](https://img.shields.io/badge/license-Apache%20License%202.0-blue.svg)](https://opensource.org/license/apache-2-0)

## 🛟 Support

For support or questions contact: 📧 [soporte@kanopus.cl](mailto:soporte@kanopus.cl)
