package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.ExcelTestSupport;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import io.github.vatisteve.utils.excel.common.ExcelElement;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Covers the stateful cursor model of {@code ExcelWriterImpl}: row/sheet positioning,
 * auto-increment, styling, output, and the error/guard conditions.
 */
@DisplayName("ExcelWriter state & positioning")
class ExcelWriterStateTest {

    @Nested
    @DisplayName("happy paths")
    class HappyPaths {

        @Test
        @DisplayName("getWorkbook() returns the backing workbook")
        void exposesWorkbook() throws IOException {
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                assertNotNull(writer.getWorkbook());
            }
        }

        @Test
        @DisplayName("cells fill left-to-right and rows advance top-to-bottom")
        void fillsCellsAndRows() throws IOException {
            byte[] bytes;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                writer.startNewRow();
                writer.addCell("a0");
                writer.addCell("a1");
                writer.startNewRow();
                writer.addCell("b0");
                bytes = writer.build();
            }
            ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
                assertEquals("a0", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("a1", sheet.getRow(0).getCell(1).getStringCellValue());
                assertEquals("b0", sheet.getRow(1).getCell(0).getStringCellValue());
            });
        }

        @Test
        @DisplayName("startNewRow(height) sets the row height")
        void setsRowHeight() throws IOException {
            byte[] bytes;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                writer.startNewRow((short) 480);
                writer.addCell("x");
                bytes = writer.build();
            }
            ExcelTestSupport.assertFirstSheet(bytes, sheet -> assertEquals((short) 480, sheet.getRow(0).getHeight()));
        }

        @Test
        @DisplayName("autoIncrementCell() emits 1,2,3 across rows")
        void autoIncrements() throws IOException {
            byte[] bytes;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                for (int i = 0; i < 3; i++) {
                    writer.startNewRow();
                    writer.autoIncrementCell();
                }
                bytes = writer.build();
            }
            ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
                assertEquals(1d, sheet.getRow(0).getCell(0).getNumericCellValue());
                assertEquals(2d, sheet.getRow(1).getCell(0).getNumericCellValue());
                assertEquals(3d, sheet.getRow(2).getCell(0).getNumericCellValue());
            });
        }

        @Test
        @DisplayName("setCellStyle styles the current cell")
        void appliesCellStyle() throws IOException {
            byte[] bytes;
            short styleIndex;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                CellStyle style = writer.getWorkbook().createCellStyle();
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                styleIndex = style.getIndex();
                writer.startNewRow();
                writer.setCellStyle(style);
                bytes = writer.build();
            }
            ExcelTestSupport.assertFirstSheet(bytes, sheet ->
                    assertEquals(styleIndex, sheet.getRow(0).getCell(0).getCellStyle().getIndex()));
        }

        @Test
        @DisplayName("build(OutputStream) writes the workbook to the stream")
        void buildsToOutputStream() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                writer.startNewRow();
                writer.addCell("streamed");
                writer.build(out);
            }
            byte[] bytes = out.toByteArray();
            assertTrue(bytes.length > 0);
            ExcelTestSupport.assertFirstSheet(bytes, sheet ->
                    assertEquals("streamed", sheet.getRow(0).getCell(0).getStringCellValue()));
        }

        @Test
        @DisplayName("startAtRow re-positions on a row written earlier in the session")
        void startsAtExistingRow() throws IOException, ElementNotFoundException {
            // Note: SXSSF can only revisit rows still inside the streaming window (i.e. rows
            // created in this session), not pre-existing rows loaded from a template.
            byte[] bytes;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                writer.startNewRow();
                writer.addCell("row0-original");
                writer.startNewRow();
                writer.addCell("row1");
                writer.startAtRow(0);          // jump back to row 0
                writer.addCell("row0-overwritten");
                bytes = writer.build();
            }
            ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
                assertEquals("row0-overwritten", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("row1", sheet.getRow(1).getCell(0).getStringCellValue());
            });
        }
    }

    @Nested
    @DisplayName("error & guard conditions")
    class ErrorConditions {

        @Test
        @DisplayName("adding a cell before starting a row throws IllegalStateException")
        void noActiveRow() throws IOException {
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                assertThrows(IllegalStateException.class, () -> writer.addCell("oops"));
            }
        }

        @Test
        @DisplayName("startAtSheet with an unknown sheet index throws ElementNotFoundException")
        void unknownSheet() throws IOException {
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                ElementNotFoundException ex =
                        assertThrows(ElementNotFoundException.class, () -> writer.startAtSheet(9, 0, 0));
                assertEquals(ExcelElement.SHEET, ex.getElement());
            }
        }

        @Test
        @DisplayName("startAtRow with a non-existent row throws ElementNotFoundException")
        void unknownRow() throws IOException, ElementNotFoundException {
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
                writer.startAtSheet(0, 0, 0);
                ElementNotFoundException ex =
                        assertThrows(ElementNotFoundException.class, () -> writer.startAtRow(5));
                assertEquals(ExcelElement.ROW, ex.getElement());
            }
        }
    }
}
