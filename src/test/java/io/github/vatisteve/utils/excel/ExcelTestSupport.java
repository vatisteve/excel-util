package io.github.vatisteve.utils.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Shared helpers for the test suite.
 * <p>
 * Verification deliberately re-reads written bytes with Apache POI <em>directly</em>
 * (via {@link WorkbookFactory}) rather than through {@code ExcelLoader}, so that a bug in
 * the loader cannot mask a bug in the writer (and vice versa). Because a POI
 * {@link Workbook}/{@link Cell} becomes unusable once closed, callers extract the values
 * they need inside the provided callback while the workbook is still open.
 */
public final class ExcelTestSupport {

    private ExcelTestSupport() {
    }

    /** Runs assertions against the first sheet of the workbook decoded from {@code bytes}. */
    public static void assertFirstSheet(byte[] bytes, Consumer<Sheet> assertions) throws IOException {
        try (Workbook wb = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertions.accept(wb.getSheetAt(0));
        }
    }

    /** Extracts a value from a single cell (first sheet) of the workbook decoded from {@code bytes}. */
    public static <T> T readCell(byte[] bytes, int row, int column, Function<Cell, T> reader) throws IOException {
        try (Workbook wb = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            Cell cell = wb.getSheetAt(0).getRow(row).getCell(column);
            return reader.apply(cell);
        }
    }
}
