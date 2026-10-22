# Excel Library

Standalone Java 21 library for annotated Excel import and export, built on Apache POI.

The code was migrated from Core Insure's Excel helper package. Its package is now
`com.ravzeno.excel`, with no Spring or Lombok dependency.

## Use from Java projects

Publish the library to your local Maven repository:

```bash
./gradlew publishToMavenLocal
```

Then add this dependency to a Gradle JVM project:

```kotlin
repositories {
    mavenLocal()
    mavenCentral()
}

dependencies {
    implementation("com.ravzeno:excel-library:1.0.0")
}
```

Use imports beginning with:

```java
com.ravzeno.excel
```

To migrate from Core Insure, replace the former package prefix
`com.x_connext.go_max_service.library.excel` with `com.ravzeno.excel`.

For imports, pass `file.getInputStream()` or `file.getBytes()` to
`AbstractExcelImportService#importFromExcel`; use `AbstractExcelExportService` for workbook
exports. The library intentionally does not depend on
Spring's `MultipartFile`.

For simple schemas, no service subclass is needed:

```java
var exporter = AbstractExcelExportService.of(
    OrderRow.class, "Orders", row -> new Object[] {row.code(), row.amount()});
var importer = AbstractExcelImportService.of(OrderImportRow.class, 1, 1, 2);
var xlsx = exporter.export(orders).bytes();
var result = importer.importFromExcel(xlsx);
```

When an `@ExcelImportColumn` supplies `headerName`, the uploaded header is checked before any
data is mapped. `ExcelImportError` coordinates are zero-based physical worksheet coordinates.

## Templates, validation, and error feedback

Create a header-only template with `exportTemplate(sheetName, filename)`, or one workbook with
several worksheets sharing the same row schema through `exportWorkbook(requests, filename)`.
Exported sheets freeze the header by default; set `freezeHeader` to `false` on
`ExcelExportRequest` when needed. Add `allowedValues` to `@ExcelColumn` to create a dropdown.

Imports support `minValue`, `maxValue`, `allowedValues`, and the existing length/pattern rules.
For domain rules involving more than one field, override `validateRow`; returning any
`ExcelImportError` rejects that row. To return a user-correctable spreadsheet, create a marked
copy of the original upload:

```java
byte[] errorWorkbook = ExcelImportErrorWorkbook.create(originalXlsx, importResult);
```

The resulting XLSX highlights invalid cells and adds each error as a cell comment.

## Large files

`AbstractExcelExportService` uses Apache POI's streaming workbook and retains only 100 generated
rows in memory. For a large export, pass a lazy `Iterable<T>` (for example, one that pages a
database query) and write directly to an HTTP response or file stream:

```java
exportService.writeTo(rows, response.getOutputStream());
```

Imports use Apache POI's SAX reader, so the workbook itself is not loaded into memory. Prefer the
`InputStream` overload for uploads. `ExcelImportResult` still retains the mapped rows and errors;
process or page extremely large imports separately if retaining every result is not acceptable.

Build and test with the included Gradle 9.3.0 wrapper:

```bash
./gradlew test
```

The wrapper allows up to two minutes for its initial download, which avoids short network
timeouts during IDE import.

## Publish to GitHub Packages

The release workflow publishes `com.ravzeno:excel-library` to GitHub Packages when a GitHub
Release is created. Published versions are immutable, so update `version` in `build.gradle.kts`
before each release.

For a local publish, create `gradle.properties` (which is ignored by Git) with a GitHub classic
personal access token that has `write:packages`:

```properties
gpr.user=YOUR_GITHUB_USERNAME
gpr.key=YOUR_GITHUB_CLASSIC_PAT
```

Then run:

```bash
./gradlew publish
```

Projects that consume this package must add the GitHub Packages Maven repository and authenticate
with a GitHub classic personal access token that has `read:packages`.
