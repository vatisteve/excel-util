package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.AbstractUtilsTest;
import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import io.github.vatisteve.utils.excel.loader.ExcelLoaderTest;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Assert;
import org.junit.Test;
import org.junit.jupiter.api.DisplayName;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.IntStream;

public class ExcelWriterTest extends AbstractUtilsTest {

    private static final Path TEMP_PATH = Paths.get(System.getProperty("java.io.tmpdir")).resolve("ExcelWriterTest");

    @Test
    @DisplayName("Basic 1")
    public void writeNewExcelData() throws IOException, ElementNotFoundException {
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
            byte[] result = writer.build();
            Assert.assertNotNull(result);
            FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-1.xlsx").toFile(), result);
        }
    }

    @Test
    @DisplayName("Basic 2")
    public void writeNewExcelWithHeader() throws IOException, ElementNotFoundException {
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
            byte[] result = writer.build();
            Assert.assertNotNull(result);
            Assert.assertTrue(result.length > 0);
            FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-2.xlsx").toFile(), result);
        }
    }

    @Test
    @DisplayName("Basic 3")
    public void writeNewExcelWithTemplate() throws IOException, ElementNotFoundException {
        InputStream templateStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        Assert.assertNotNull(templateStream);
        byte[] inByteArray = IOUtils.toByteArray(templateStream);
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new ByteArrayInputStream(inByteArray))) {
            writer.startAtSheet(0, 701, 0);
            writer.startNewRow();
            SampleDomain sample = ExcelLoaderTest.getSampleDomainData(new ByteArrayInputStream(inByteArray));
            sample.writeDataTo(writer);
            byte[] result = writer.build();
            FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-3.xlsx").toFile(), result);
        }
    }

    @Test
    @DisplayName("Merge cell with Cell Operation Function interface")
    public void writeNewExcelWithCellOperationFunction() throws IOException, ElementNotFoundException {
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
            byte[] result = writer.build();
            Assert.assertNotNull(result);
            FileUtils.writeByteArrayToFile(TEMP_PATH.resolve("output-4.xlsx").toFile(), result);
        }
    }

}
