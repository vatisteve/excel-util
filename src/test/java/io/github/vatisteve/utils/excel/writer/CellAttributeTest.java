package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;

@DisplayName("CellAttribute builder")
class CellAttributeTest {

    @Test
    @DisplayName("no-arg builder leaves the style null")
    void withoutStyle() {
        CellAttribute attribute = new CellAttribute.Builder().value("v").build();
        assertNull(attribute.getCellStyle());
        assertEquals("v", attribute.getValue());
        assertNull(attribute.getCellOperation());
    }

    @Test
    @DisplayName("builder retains style, value and operation")
    void withAllFields() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            CellStyle style = wb.createCellStyle();
            CellOperation operation = (sheet, cell) -> cell;
            CellAttribute attribute = new CellAttribute.Builder(style)
                    .value(7)
                    .cellOperation(operation)
                    .build();
            assertSame(style, attribute.getCellStyle());
            assertEquals(7, attribute.getValue());
            assertSame(operation, attribute.getCellOperation());
        }
    }
}
