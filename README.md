# Excel Library

**A lightweight Java library for importing and exporting Excel workbooks with annotated models.**

Built on Apache POI, this standalone Java 21 library provides XLSX import and export, field-level
validation, reusable templates, and streaming APIs for large files. It has no Spring or Lombok
dependency, so it can be used in any Java application.

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick start](#quick-start)
- [Import validation](#import-validation)
- [Templates and multiple sheets](#templates-and-multiple-sheets)
- [Streaming large imports and exports](#streaming-large-imports-and-exports)
- [Date handling](#date-handling)
- [CSV imports](#csv-imports)
- [Spring Boot integration](#spring-boot-integration)
- [Publishing](#publishing)

## Features

- Map annotated Java fields to worksheet columns.
- Import XLSX files with required-field, length, numeric-range, allowed-value, and pattern checks.
- Validate uploaded headers and apply custom rules across fields.
- Export annotated rows, header-only templates, and workbooks with multiple sheets.
- Create a marked copy of an uploaded workbook with invalid cells highlighted and annotated.
- Stream XLSX rows with Apache POI's SAX reader and export rows with a streaming workbook.
- Import CSV files using the same annotations and validation rules.

## Requirements

- Java 21 or later
- Gradle for building from source

## Installation

### Use the published package

Add the GitHub Packages repository and dependency to your Gradle build. GitHub Packages requires
authentication with a classic personal access token that has `read:packages` permission.

```kotlin
repositories {
    maven {
        url = uri("https://maven.pkg.github.com/ravzeno/excel-library")
        credentials {
            username = providers.gradleProperty("gpr.user").orNull
            password = providers.gradleProperty("gpr.key").orNull
        }
    }
    mavenCentral()
}

dependencies {
    implementation("com.ravzeno:excel-library:1.0.0")
}
```

Keep credentials in your local `~/.gradle/gradle.properties` or another secure secret store:

```properties
gpr.user=YOUR_GITHUB_USERNAME
gpr.key=YOUR_GITHUB_CLASSIC_PAT
```

### Build from source

```bash
git clone https://github.com/ravzeno/excel-library.git
cd excel-library
./gradlew build
```

To install the artifact in your local Maven repository:

```bash
./gradlew publishToMavenLocal
```

## Quick start

Import and export models use separate annotations. Column indexes are zero-based.

```java
public class OrderRow {
    @ExcelColumn(name = "Order code", index = 0)
    private String code;

    @ExcelColumn(name = "Amount", index = 1)
    private BigDecimal amount;

    // Getters and constructor omitted
}

public class OrderImportRow {
    @ExcelImportColumn(index = 0, headerName = "Order code", required = true)
    private String code;

    @ExcelImportColumn(index = 1, headerName = "Amount", dataType = ExcelDataType.NUMERIC)
    private BigDecimal amount;

    public OrderImportRow() {}
}
```

Create services without writing dedicated subclasses. The exporter reads values from fields marked
with `@ExcelColumn`, ordered by the annotation index:

```java
var exporter = AbstractExcelExportService.of(OrderRow.class, "Orders");
var importer = AbstractExcelImportService.of(OrderImportRow.class, 1, 1, 2);

byte[] xlsx = exporter.export(orders).bytes();
ExcelImportResult<OrderImportRow> result = importer.importFromExcel(xlsx);

System.out.printf("Imported %d of %d rows%n", result.validRows(), result.totalRows());
```

When export values need transformation, use the mapper overload or override `mapToRow` in a
subclass:

```java
var exporter = AbstractExcelExportService.of(
    OrderRow.class, "Orders",
    row -> new Object[] {row.getCode(), row.getAmount().setScale(2)}
);
```

The importer factory arguments are the entity class, one-based worksheet number, one-based header
row, and one-based first data row. For upload streams, call `importFromExcel(inputStream)`; the
stream is consumed but not closed by the library.

## Import validation

`@ExcelImportColumn` supports required values, text length limits, numeric limits, enumerated
values, and regular-expression patterns:

```java
@ExcelImportColumn(
    index = 0,
    headerName = "Status",
    required = true,
    allowedValues = {"NEW", "PAID", "CANCELLED"}
)
private String status;
```

When `headerName` is set, the uploaded header is checked before any data rows are mapped. Import
results contain valid rows, counts, and structured errors. Error row and column coordinates are
zero-based physical worksheet coordinates.

For rules involving more than one field, override `validateRow` in an importer subclass. Return one
or more `ExcelImportError` values to reject that row. To give users a corrected copy of their
upload, create an error workbook:

```java
byte[] errorWorkbook = ExcelImportErrorWorkbook.create(originalXlsx, result);
```

The error workbook highlights invalid cells and attaches each message as a cell comment.

## Templates and multiple sheets

Create a header-only template with `exportTemplate`. Create a workbook with several worksheets by
passing `ExcelExportRequest` values to `exportWorkbook`:

```java
var template = exporter.exportTemplate("Orders", "orders-template.xlsx");

var workbook = exporter.exportWorkbook(
    List.of(
        new ExcelExportRequest<>(pendingOrders, "Pending", "orders.xlsx", null, null),
        new ExcelExportRequest<>(completedOrders, "Completed", "orders.xlsx", null, null)
    ),
    "orders.xlsx"
);
```

Exported sheets freeze the header row by default. Set `freezeHeader` to `false` on an
`ExcelExportRequest` to disable it. `@ExcelColumn(allowedValues = ...)` adds a dropdown list to the
exported column.

## Streaming large imports and exports

The XLSX importer uses Apache POI's SAX reader. The exporter uses a streaming workbook and keeps a
bounded row window in memory. For exports, pass a lazy `Iterable` and write directly to an output
stream:

```java
exporter.writeTo(databaseRows, outputStream);
```

For imports, use `importEach` to handle each valid row immediately, or `importInBatches` to retain
only a bounded batch of mapped rows:

```java
ExcelImportSummary summary = importer.importInBatches(
    inputStream,
    500,
    batch -> orderRepository.saveAll(batch)
);
```

`ExcelImportSummary` reports row counts and validation errors without retaining the successfully
mapped rows. Errors are retained in the summary, so files with a very large number of invalid rows
can still require substantial memory.

## Date handling

Date fields support the standard `ExcelDateFormat` values. Supply a Java time pattern with
`datePattern` when a column uses a different format:

```java
@ExcelImportColumn(index = 2, datePattern = "dd-MM-yyyy")
private LocalDate orderDate;
```

Numeric date fields accept padded and unpadded days and months, such as `01/11/2026` and
`1/11/2026`. The selected date format determines whether the value is interpreted day-first or
month-first.

For fields typed as `Instant`, override `getZoneId()` in your importer to define the timezone used
to interpret imported local date-times. The default is the JVM timezone.

## CSV imports

CSV files can be imported with the same annotations, header checks, conversion, and validation as
XLSX:

```java
ExcelImportResult<OrderImportRow> result = importer.importFromCsv(csvInputStream);
```

The default character set is UTF-8. To specify another character set or worksheet-style layout,
use `importFromCsv(inputStream, charset, request)`. CSV input is converted to an intermediate
workbook before mapping, so use the XLSX streaming APIs for very large files.

## Spring Boot integration

The library does not depend on Spring. In a Spring Boot controller, pass the uploaded file's
stream directly to the importer:

```java
@PostMapping("/orders/import")
public ExcelImportResult<OrderImportRow> importOrders(
        @RequestParam("file") MultipartFile file) throws IOException {
    return orderImporter.importFromExcel(file.getInputStream());
}
```

## Publishing

GitHub Releases trigger the repository's release workflow to publish the package to GitHub
Packages. Published versions are immutable; update `version` in `build.gradle.kts` before each
release.

For a local publish, provide `gpr.user` and `gpr.key` in the repository's ignored
`gradle.properties` file, then run:

```bash
./gradlew publish
```
