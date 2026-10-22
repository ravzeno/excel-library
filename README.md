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
