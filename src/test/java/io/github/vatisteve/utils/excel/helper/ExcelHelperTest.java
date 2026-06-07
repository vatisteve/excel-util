package io.github.vatisteve.utils.excel.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

@DisplayName("ExcelHelper")
class ExcelHelperTest {

    private Workbook wb;
    private Sheet sheet;

    @BeforeEach
    void setUp() {
        wb = new XSSFWorkbook();
        sheet = wb.createSheet();
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(12.5);
        row.createCell(1).setCellValue("hi");
        row.createCell(2).setCellValue(true);
        row.createCell(3).setCellErrorValue(FormulaError.DIV0.getCode());
        Cell formula = row.createCell(4);
        formula.setCellFormula("1+2");
        wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(formula);
        sheet.createRow(1); // an empty row (no cells)
    }

    @AfterEach
    void tearDown() throws IOException {
        wb.close();
    }

    @Test
    @DisplayName("getCellValue maps each cell type to the expected Java type")
    void getCellValueByType() {
        Row row = sheet.getRow(0);
        double numeric = ExcelHelper.<Double>getCellValue(row.getCell(0));
        assertEquals(12.5d, numeric);
        assertEquals("hi", ExcelHelper.<String>getCellValue(row.getCell(1)));
        assertTrue(ExcelHelper.<Boolean>getCellValue(row.getCell(2)));
        assertNull(ExcelHelper.<Object>getCellValue(row.getCell(3))); // ERROR -> null
    }

    @Test
    @DisplayName("getCellValue resolves a formula via its cached result type")
    void getCellValueFormula() {
        double cached = ExcelHelper.<Double>getCellValue(sheet.getRow(0).getCell(4));
        assertEquals(3d, cached);
    }

    @Test
    @DisplayName("getCell returns the cell, or null for missing rows/cells")
    void getCell() {
        assertNotNull(ExcelHelper.getCell(sheet, 0, 0));   // existing
        assertNull(ExcelHelper.getCell(sheet, 0, 5));      // missing row
        assertNull(ExcelHelper.getCell(sheet, 0, 1));      // existing row, blank cell
    }
}
