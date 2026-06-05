package io.github.vatisteve.utils.excel.loader;

import io.github.vatisteve.utils.excel.AbstractUtilsTest;
import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import io.github.vatisteve.utils.excel.common.ElementIdentifier;
import io.github.vatisteve.utils.excel.common.ExcelElement;
import org.apache.poi.ss.util.CellAddress;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Exercises the full {@code ExcelLoader} surface (sheet management, the typed-getter overloads,
 * and the error/cast failure paths) against the {@code Financial_Sample.xlsx} fixture.
 * <p>
 * Reference cells in that fixture: {@code A1} = "Segment", {@code B2} = "Canada",
 * {@code F2} = 3 (numeric). The single sheet is named "Sheet1".
 */
@DisplayName("ExcelLoader API")
class ExcelLoaderApiTest extends AbstractUtilsTest {

    private static final CellAddress B2 = new CellAddress("B2");
    private static final CellAddress F2 = new CellAddress("F2");
    private static final CellAddress A1 = new CellAddress("A1");

    private static InputStream sample() {
        return ExcelLoaderApiTest.class.getResourceAsStream("/Financial_Sample.xlsx");
    }

    @Nested
    @DisplayName("sheet management")
    class SheetManagement {

        @Test
        @DisplayName("default sheet can be selected by index and by name")
        void selectDefaultSheet() throws IOException, ElementNotFoundException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertNotNull(loader.getDefaultSheet());
                loader.setDefaultSheet(0);
                assertEquals("Sheet1", loader.getDefaultSheet().getSheetName());
                loader.setDefaultSheet("Sheet1");
                assertEquals("Sheet1", loader.getDefaultSheet().getSheetName());
            }
        }

        @Test
        @DisplayName("sheet name and index lookups")
        void nameAndIndexLookups() throws IOException, ElementNotFoundException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertEquals("Sheet1", loader.getSheetName(0));
                assertEquals(0, loader.getSheetIndex("Sheet1"));
                assertEquals(-1, loader.getSheetIndex("Does Not Exist"));
            }
        }

        @Test
        @DisplayName("unknown sheet index reports SHEET-POSITION")
        void unknownSheetIndex() throws IOException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                ElementNotFoundException ex =
                        assertThrows(ElementNotFoundException.class, () -> loader.setDefaultSheet(99));
                assertEquals(ExcelElement.SHEET, ex.getElement());
                assertEquals(ElementIdentifier.POSITION, ex.getIdentifier());
            }
        }

        @Test
        @DisplayName("unknown sheet name reports SHEET-NAME")
        void unknownSheetName() throws IOException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                ElementNotFoundException ex =
                        assertThrows(ElementNotFoundException.class, () -> loader.setDefaultSheet("Nope"));
                assertEquals(ExcelElement.SHEET, ex.getElement());
                assertEquals(ElementIdentifier.NAME, ex.getIdentifier());
            }
        }
    }

    @Nested
    @DisplayName("typed getters")
    class TypedGetters {

        @Test
        @DisplayName("getString across all addressing overloads")
        void getStringOverloads() throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertEquals("Canada", loader.getString(B2));            // default sheet, address
                assertEquals("Canada", loader.getString(1, 1));         // default sheet, col/row
                assertEquals("Canada", loader.getString(0, B2));        // sheet index, address
                assertEquals("Canada", loader.getString("Sheet1", B2)); // sheet name, address
                assertEquals("Canada", loader.getString(0, 1, 1));      // sheet index, col/row
                assertEquals("Canada", loader.getString("Sheet1", 1, 1)); // sheet name, col/row
            }
        }

        @Test
        @DisplayName("getLong and getInteger across overloads")
        void numericOverloads() throws IOException, ElementNotFoundException, CastCellValueExcelLoaderException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertEquals(3L, loader.getLong(F2));
                assertEquals(3L, loader.getLong(5, 1));
                assertEquals(3L, loader.getLong(0, F2));
                assertEquals(3L, loader.getLong("Sheet1", F2));
                assertEquals(3, loader.getInteger(F2));
                assertEquals(3, loader.getInteger(0, F2));
                assertEquals(3, loader.getInteger("Sheet1", 5, 1));
            }
        }

        @Test
        @DisplayName("getValue returns the raw typed value")
        void getValueOverloads() throws IOException, ElementNotFoundException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertEquals("Canada", loader.getValue(B2));
                assertEquals("Canada", loader.getValue(1, 1));
                assertEquals("Canada", loader.getValue(0, B2));
                assertEquals("Canada", loader.getValue("Sheet1", B2));
                assertEquals(3d, (Double) loader.getValue(F2));
            }
        }
    }

    @Nested
    @DisplayName("error & cast-failure paths")
    class Errors {

        @Test
        @DisplayName("reading a missing cell throws ElementNotFoundException(CELL)")
        void missingCell() throws IOException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                ElementNotFoundException ex =
                        assertThrows(ElementNotFoundException.class, () -> loader.getString(100, 0));
                assertEquals(ExcelElement.CELL, ex.getElement());
            }
        }

        @Test
        @DisplayName("reading a String cell as a number throws CastCellValueExcelLoaderException")
        void castFailure() throws IOException {
            try (ExcelLoader loader = ExcelUtilsFactory.createExcelLoader(sample())) {
                assertThrows(CastCellValueExcelLoaderException.class, () -> loader.getLong(A1));
                assertThrows(CastCellValueExcelLoaderException.class, () -> loader.getInteger(A1));
            }
        }
    }
}
