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
import org.junit.Assert;
import org.junit.Test;
import org.junit.jupiter.api.DisplayName;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.stream.IntStream;

public class ExcelWriterTest extends AbstractUtilsTest {

    @Test
    @DisplayName("Basic 1")
    public void writeNewExcelData() throws IOException, ElementNotFoundException, ExcelWriterException {
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
            FileUtils.writeByteArrayToFile(FileUtils.getFile("E:\\Temporary\\test1.xlsx"), result);
        }
    }

    @Test
    @DisplayName("Basic 2")
    public void writeNewExcelWithHeader() throws IOException, ElementNotFoundException, ExcelWriterException {
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
            FileUtils.writeByteArrayToFile(FileUtils.getFile("E:\\Temporary\\test2.xlsx"), result);
        }
    }

    @Test
    @DisplayName("Basic 3")
    public void writeNewExcelWithTemplate() throws IOException, ExcelWriterException, ElementNotFoundException {
        InputStream templateStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        Assert.assertNotNull(templateStream);
        byte[] inByteArray = IOUtils.toByteArray(templateStream);
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new ByteArrayInputStream(inByteArray))) {
            writer.startAtSheet(0, 701, 0);
            writer.startNewRow();
            SampleDomain sample = ExcelLoaderTest.getSampleDomainData(new ByteArrayInputStream(inByteArray));
            sample.writeDataTo(writer);
            byte[] result = writer.build();
            FileUtils.writeByteArrayToFile(FileUtils.getFile("E:\\Temporary\\test3.xlsx"), result);
        }
    }
}
