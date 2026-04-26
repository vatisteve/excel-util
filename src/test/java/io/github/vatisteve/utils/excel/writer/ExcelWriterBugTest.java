package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.AbstractUtilsTest;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.time.ZonedDateTime;
import java.time.OffsetDateTime;

public class ExcelWriterBugTest extends AbstractUtilsTest {

    /**
     * This test was created to verify a fix for a bug where custom {@link CellStyle} was ignored
     * when adding {@link ZonedDateTime} or {@link OffsetDateTime} values to a cell.
     * It also verifies that the generic {@code addCell(Object, CellStyle)} method correctly
     * applies the style for these specific date-time types.
     */
    @Test
    public void testStyleBugInAddCell() throws IOException {
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            writer.startNewRow();
            
            CellStyle redStyle = writer.getWorkbook().createCellStyle();
            redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            ZonedDateTime nowZoned = ZonedDateTime.now();
            OffsetDateTime nowOffset = OffsetDateTime.now();
            
            writer.addCell(nowZoned, redStyle);
            writer.addCell(nowOffset, redStyle);
            
            Sheet sheet = writer.getWorkbook().getSheetAt(0);
            Cell cell1 = sheet.getRow(0).getCell(0);
            Cell cell2 = sheet.getRow(0).getCell(1);
            
            Assert.assertEquals("ZonedDateTime cell should have red style", redStyle.getIndex(), cell1.getCellStyle().getIndex());
            Assert.assertEquals("OffsetDateTime cell should have red style", redStyle.getIndex(), cell2.getCellStyle().getIndex());
        }
    }
}
