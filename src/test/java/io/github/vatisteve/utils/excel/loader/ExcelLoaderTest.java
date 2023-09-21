package io.github.vatisteve.utils.excel.loader;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.util.CellAddress;
import org.junit.Test;
import org.junit.jupiter.api.DisplayName;

import io.github.vatisteve.utils.excel.ExcelUtilsFactory;

public class ExcelLoaderTest {

    @Test
    @DisplayName("New Excel loader from local file")
    public void newInstance() {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals("Segment", excelLoader.getString(new CellAddress("A1")));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ElementExcelLoaderNotFoundException e) {
            e.printStackTrace();
        } catch (CastCellValueExcelLoaderException e) {
            e.printStackTrace();
        }
    }

    @Test
    @DisplayName("Check string type")
    public void checkStringType() {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals("Canada", excelLoader.getString(new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getString(0, new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getString("Sheet1", new CellAddress("B2")));
            assertEquals("Canada", excelLoader.getValue(1, 1));
            assertEquals("Canada", excelLoader.getValue(0, 1, 1));
            assertEquals("Canada", excelLoader.getValue("Sheet1", 1, 1));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ElementExcelLoaderNotFoundException e) {
            e.printStackTrace();
        } catch (CastCellValueExcelLoaderException e) {
            e.printStackTrace();
        }
    }

    @Test
    @DisplayName("Check numberic type")
    public void checkNumbericType() {
        InputStream sourceStream = this.getClass().getResourceAsStream("/Financial_Sample.xlsx");
        try(ExcelLoader excelLoader = ExcelUtilsFactory.createExcelLoader(sourceStream)) {
            assertEquals(3, excelLoader.getLong(new CellAddress("F2")));
            assertEquals(3, excelLoader.getInteger(new CellAddress("F2")));
            assertEquals(3, ((Double) excelLoader.getValue(new CellAddress("F2"))).intValue());
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ElementExcelLoaderNotFoundException e) {
            e.printStackTrace();
        } catch (CastCellValueExcelLoaderException e) {
            e.printStackTrace();
        }
    }
}