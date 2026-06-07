package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ExcelTestSupport;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Focused tests for {@code addCell(CellAttribute)} and its {@link CellOperation} hook,
 * including the regression guard for the previously-broken failure path (a thrown
 * {@code ExcelWriterException} used to skip a column and silently drop the value/style).
 */
@DisplayName("ExcelWriter cell operations")
class ExcelWriterCellOperationTest {

    @Test
    @DisplayName("the value is written to the cell returned by the operation")
    void operationCanRedirectTargetCell() throws IOException {
        byte[] bytes;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            writer.startNewRow();
            writer.addCell("a");   // col 0
            writer.addCell("b");   // col 1
            // The operation creates the cell at col 2 but redirects the value to col 5.
            writer.addCell(new CellAttribute.Builder()
                    .value("redirected")
                    .cellOperation((sheet, cell) -> cell.getRow().createCell(5))
                    .build());
            writer.addCell("after"); // col 3
            bytes = writer.build();
        }
        ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
            assertEquals("redirected", sheet.getRow(0).getCell(5).getStringCellValue());
            assertEquals("after", sheet.getRow(0).getCell(3).getStringCellValue());
        });
    }

    @Test
    @DisplayName("a failing operation falls back to the plain cell without skipping a column or losing the style")
    void failingOperationFallsBackCleanly() throws IOException {
        byte[] bytes;
        short styleIndex;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            CellStyle style = writer.getWorkbook().createCellStyle();
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            styleIndex = style.getIndex();

            writer.startNewRow();
            writer.addCell(new CellAttribute.Builder(style)
                    .value("V")
                    .cellOperation((sheet, cell) -> {
                        throw new ExcelWriterException("boom");
                    })
                    .build());        // must land in col 0, keep its style
            writer.addCell("next");   // must land in col 1 (no skipped column)
            bytes = writer.build();
        }
        short expectedStyle = styleIndex;
        ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
            assertEquals("V", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(expectedStyle, sheet.getRow(0).getCell(0).getCellStyle().getIndex());
            assertEquals("next", sheet.getRow(0).getCell(1).getStringCellValue());
            // col 1 should be the very next cell — confirm nothing was orphaned in between
            assertEquals(CellType.STRING, sheet.getRow(0).getCell(1).getCellType());
        });
    }
}
