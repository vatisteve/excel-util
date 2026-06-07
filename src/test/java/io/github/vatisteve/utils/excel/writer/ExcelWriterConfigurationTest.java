package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ExcelTestSupport;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import io.github.vatisteve.utils.excel.writer.ExcelWriterConfiguration.DefaultConfiguration;
import io.github.vatisteve.utils.excel.writer.ExcelWriterConfiguration.ExcelHeader;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.time.LocalTime;
import java.time.ZoneId;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

@DisplayName("ExcelWriterConfiguration")
class ExcelWriterConfigurationTest {

    @Nested
    @DisplayName("DefaultConfiguration")
    class Defaults {

        private final ExcelWriterConfiguration config = new DefaultConfiguration();

        @Test
        @DisplayName("provides documented default values")
        void defaultValues() throws IOException {
            assertEquals("Data 0", config.sheetName(0));
            assertEquals("HH:mm:ss", config.timeFormat());
            assertEquals(ZoneId.systemDefault(), config.zoneId());
            assertEquals((short) -1, config.rowHeight());
            try (Workbook wb = new XSSFWorkbook()) {
                assertNotNull(config.cellStyle(wb));
                assertNull(config.excelHeader(wb));
            }
        }
    }

    @Nested
    @DisplayName("ExcelHeader builder")
    class Header {

        @Test
        @DisplayName("captures headers, height and sheet index")
        void buildsHeader() {
            ExcelHeader header = new ExcelHeader.Builder()
                    .headers("A", "B")
                    .height((short) 300)
                    .sheetIndex(2)
                    .build();
            assertArrayEquals(new String[]{"A", "B"}, header.getHeaders());
            assertEquals((short) 300, header.getHeight());
            assertEquals(2, header.getSheetIndex());
        }

        @Test
        @DisplayName("defaults height to -1 when unset")
        void defaultHeight() {
            assertEquals((short) -1, new ExcelHeader.Builder().headers("X").build().getHeight());
        }
    }

    @Nested
    @DisplayName("custom configuration is applied to the workbook")
    class CustomConfig {

        @Test
        @DisplayName("sheet name, header row, row height and time format all take effect")
        void appliesCustomConfig() throws IOException {
            ExcelWriterConfiguration config = new ExcelWriterConfiguration() {
                @Override
                public String sheetName(int index) {
                    return "Custom";
                }

                @Override
                public ExcelHeader excelHeader(Workbook wb) {
                    return new ExcelHeader.Builder().headers("H1", "H2").build();
                }

                @Override
                public short rowHeight() {
                    return 360;
                }

                @Override
                public String timeFormat() {
                    return "HH:mm";
                }
            };

            byte[] bytes;
            try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(config)) {
                // header was written at row 0 by the constructor; data starts at row 1
                writer.startNewRow();
                writer.addCell(LocalTime.of(9, 5, 0));
                bytes = writer.build();
            }

            ExcelTestSupport.assertFirstSheet(bytes, sheet -> {
                assertEquals("Custom", sheet.getSheetName());
                assertEquals("H1", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("H2", sheet.getRow(0).getCell(1).getStringCellValue());
                assertEquals("09:05", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals((short) 360, sheet.getRow(1).getHeight());
            });
        }
    }
}
