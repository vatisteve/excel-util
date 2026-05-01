# excel-util

A lightweight Java utility library for reading, creating, and updating Excel files using Apache POI.

## Features

- **Reading**: Extract values from `InputStream` or existing files using Excel addresses (e.g., `A1`) or indices.
- **Writing**: Create new workbooks or update existing templates with a simple row-by-row API.
- **Styling**: Apply custom styles, headers, and row heights via simple configurations.
- **Advanced**: Supports auto-incrementing cells and custom operations like merging.

## Installation

Add the following dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>io.github.vatisteve</groupId>
    <artifactId>excel-util</artifactId>
    <version>1.0.0</version>
</dependency>
```

## Quick Start

### Read Excel

```java
try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(inputStream)) {
    String value = loader.getString(new CellAddress("A1"));
    int units = loader.getInteger(new CellAddress("F2"));
}
```

### Write Excel

```java
try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
    writer.startNewRow();
    writer.addCell("ID");
    writer.addCell("Name");

    writer.startNewRow();
    writer.autoIncrementCell(); // Auto-increment ID
    writer.addCell("Product A");

    byte[] result = writer.build(); // Export to byte array or OutputStream
}
```

## Advanced Usage

### Custom Configuration
Implement `ExcelWriterConfiguration` to define global styles, headers, or time zones:

```java
public class MyConfig implements ExcelWriterConfiguration {
    @Override
    public CellStyle cellStyle(Workbook wb) { /* Custom default style */ }

    @Override
    public ExcelHeader excelHeader(Workbook wb) {
        return new ExcelHeader.Builder().headers("H1", "H2").build();
    }
}
// Usage
ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new MyConfig());
```

### Template & Merging
```java
// Load from template
ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(templateInputStream);

// Custom operation (e.g., Merge cells)
writer.addCell(new CellAttribute.Builder()
    .value("Merged")
    .cellOperation((sheet, cell) -> {
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        return cell;
    }).build());
```

## License

See [`LICENSE`](LICENSE).
