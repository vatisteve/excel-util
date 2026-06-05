package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.AbstractUtilsTest;
import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import io.github.vatisteve.utils.excel.loader.CastCellValueExcelLoaderException;
import io.github.vatisteve.utils.excel.loader.ExcelLoader;
import io.github.vatisteve.utils.excel.loader.ExcelLoaderTest;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelWriterTest extends AbstractUtilsTest {

    private static final Path TEMP_PATH = Paths.get(System.getProperty("java.io.tmpdir")).resolve("ExcelWriterTest");

    @Test
    @DisplayName("Basic 1")
    public void writeNewExcelData() throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
        byte[] result;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            CellStyle defaultCellStyle = writer.getWorkbook().createCellStyle();
            defaultCellStyle.setAlignment(HorizontalAlignment.CENTER);
            defaultCellStyle.setBorderBottom(BorderStyle.MEDIUM_DASHED);
            defaultCellStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
            writer.startAtSheet(0, 0, 0);
            IntStream.range(1, 100)
                    .mapToObj(i -> new CellAttribute.Builder(defaultCellStyle).value("The number " + i).build())
                    .forEach(data -> {
                        writer.startNewRow();
                        writer.autoIncrementCell();
                        writer.addCell(data);
                    });
            result = writer.build();
        }
        assertNotNull(result);
        FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-1.xlsx").toFile(), result);

        // round-trip: the written workbook must read back the same values
        try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(new ByteArrayInputStream(result))) {
            assertEquals(1, loader.getInteger(0, 0));
            assertEquals("The number 1", loader.getString(1, 0));
            assertEquals(99, loader.getInteger(0, 98));
            assertEquals("The number 99", loader.getString(1, 98));
        }
    }

    @Test
    @DisplayName("Basic 2")
    public void writeNewExcelWithHeader() throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
        byte[] result;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new SampleExcelWriterConfig())) {
            writer.startAtSheet(0, 1, 0);
            IntStream.range(1, 100).mapToObj(i -> "R" + i)
                    .forEach(row -> {
                        writer.startNewRow();
                        writer.autoIncrementCell();
                        IntStream.of(1, 2, 3).mapToObj(i -> "C" + i)
                                .map(column -> String.format("[%s,%s]", row, column))
                                .forEach(writer::addCell);
                    });
            result = writer.build();
        }
        assertNotNull(result);
        assertTrue(result.length > 0);
        FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-2.xlsx").toFile(), result);

        // round-trip: header occupies row 0, data starts at row 1
        try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(new ByteArrayInputStream(result))) {
            assertEquals("Header 1", loader.getString(0, 0));
            assertEquals("And more", loader.getString(3, 0));
            assertEquals(1, loader.getInteger(0, 1));
            assertEquals("[R1,C1]", loader.getString(1, 1));
            assertEquals("[R99,C3]", loader.getString(3, 99));
        }
    }

    @Test
    @DisplayName("Basic 3")
    public void writeNewExcelWithTemplate()
            throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
        InputStream templateStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        assertNotNull(templateStream);
        byte[] inByteArray = IOUtils.toByteArray(templateStream);
        byte[] result;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new ByteArrayInputStream(inByteArray))) {
            writer.startAtSheet(0, 701, 0);
            writer.startNewRow();
            SampleDomain sample = ExcelLoaderTest.getSampleDomainData(new ByteArrayInputStream(inByteArray));
            sample.writeDataTo(writer);
            result = writer.build();
        }
        assertNotNull(result);
        assertTrue(result.length > 0);
        FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-3.xlsx").toFile(), result);

        // round-trip: original template content is preserved and the new row was appended
        try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(new ByteArrayInputStream(result))) {
            assertEquals("Segment", loader.getString(new CellAddress("A1")));
            // the appended row carries the segment value read from the template's first data row
            assertEquals("Government", loader.getString(0, 701));
        }
    }

    @Test
    @DisplayName("Merge cell with Cell Operation Function interface")
    public void writeNewExcelWithCellOperationFunction()
            throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
        byte[] result;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            writer.startAtSheet(0, 0, 0);
            // Row 1
            writer.startNewRow();
            writer.autoIncrementCell();
            writer.addCell(1);
            writer.addCell(2);
            writer.addCell(3);
            writer.addCell(4);
            // Row 2
            writer.startNewRow();
            writer.autoIncrementCell();
            // merge above
            writer.addCell(new CellAttribute.Builder().value("Merge above")
                            .cellOperation((sheet, cell) -> {
                                int currentRow = cell.getRowIndex();
                                int currentCol = cell.getColumnIndex();
                                CellRangeAddress mergeRange = new CellRangeAddress(currentRow - 1, currentRow, currentCol, currentCol);
                                sheet.addMergedRegion(mergeRange);
                                // return the first (top-left) cell of merged range
                                return sheet.getRow(currentRow - 1).getCell(currentCol);
                            })
                            .build());
            writer.addCell("After merging");
            writer.addCell("Will be merged");
            // merge previous
            writer.addCell(new CellAttribute.Builder().value("Merge previous")
                            .cellOperation((sheet, cell) -> {
                                int currentRow = cell.getRowIndex();
                                int currentCol = cell.getColumnIndex();
                                CellRangeAddress mergeRange = new CellRangeAddress(currentRow, currentRow, currentCol - 1, currentCol);
                                sheet.addMergedRegion(mergeRange);
                                return sheet.getRow(currentRow).getCell(currentCol - 1);
                            })
                            .build());
            writer.addCell("After another merge");
            // and more
            // ...
            result = writer.build();
        }
        assertNotNull(result);
        FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-4.xlsx").toFile(), result);

        // round-trip: both merged regions and the merged values must be present, and no
        // column was skipped by the cell-operation path
        try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(new ByteArrayInputStream(result))) {
            assertEquals(2, loader.getDefaultSheet().getNumMergedRegions());
            assertEquals("Merge above", loader.getString(1, 0));
            assertEquals("After merging", loader.getString(2, 1));
            assertEquals("Merge previous", loader.getString(3, 1));
            assertEquals("After another merge", loader.getString(5, 1));
        }
    }

}
