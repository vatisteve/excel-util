package io.github.vatisteve.utils.excel.loader;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.Date;

import io.github.vatisteve.utils.excel.AbstractUtilsTest;
import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.writer.SampleDomain;
import org.apache.poi.ss.util.CellAddress;
import org.junit.Assert;
import org.junit.Test;
import org.junit.jupiter.api.DisplayName;

import io.github.vatisteve.utils.excel.ExcelUtilsFactory;

public class ExcelLoaderTest extends AbstractUtilsTest {

    @Test
    @DisplayName("New Excel loader from local file")
    public void newInstance() throws IOException, CastCellValueExcelLoaderException, ElementNotFoundException {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals("Segment", excelLoader.getString(new CellAddress("A1")));
        }
    }

    @Test
    @DisplayName("Check string type")
    public void checkStringType() throws IOException, CastCellValueExcelLoaderException, ElementNotFoundException {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals("Canada", excelLoader.getString(new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getString(0, new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getString("Sheet1", new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getValue(1, 1));
            assertEquals("Canada", excelLoader.getValue(0, 1, 1));
            assertEquals("Canada", excelLoader.getValue("Sheet1", 1, 1));
        }
    }

    @Test
    @DisplayName("Check numeric type")
    public void checkNumericType() throws IOException, CastCellValueExcelLoaderException, ElementNotFoundException {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals(3, excelLoader.getLong(new CellAddress("F2")));
            assertEquals(3, excelLoader.getInteger(new CellAddress("F2")));
            assertEquals(3, ((Double) excelLoader.getValue(new CellAddress("F2"))).intValue());
        }
    }

    public static SampleDomain getSampleDomainData(InputStream sourceStream) {
        SampleDomain sample = new SampleDomain();
        try(ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            sample.setSegment(loader.getString(0, 1));
            sample.setCountry(loader.getString(new CellAddress("B4")));
            sample.setProduct(loader.getString(2, 19));
            sample.setDiscountBand(loader.getString(3, 1));
            sample.setUnitsSold(loader.getValue(new CellAddress("E2")));
            sample.setGrossSales(BigDecimal.valueOf((Double) loader.getValue(new CellAddress("H14"))));
            sample.setDate(LocalDate.of(1900, Month.JANUARY, 1).plusDays(loader.getLong(new CellAddress("M2")) - 2));
        } catch (IOException | ElementNotFoundException | CastCellValueExcelLoaderException ignored) {}
        return sample;
    }
}
