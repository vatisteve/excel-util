package io.github.vatisteve.utils.excel.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

public final class ExcelHelper {

    private ExcelHelper() {
    }

    public static <T> T getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            return getCellValue(cell, cell.getCachedFormulaResultType());
        } else if (cellType == CellType._NONE || cellType == CellType.ERROR) {
            return null;
        } else {
            return getCellValue(cell, cellType);
        }
    }

    @SuppressWarnings("unchecked")
    public static <T> T getCellValue(Cell cell, CellType cellType) {
        switch (cellType) {
        case BOOLEAN:
            return (T) Boolean.valueOf(cell.getBooleanCellValue());
        case NUMERIC:
            return (T) Double.valueOf(cell.getNumericCellValue());
        case STRING:
            return (T) cell.getStringCellValue();
        default:
            return null;
        }
    }

    public static Cell getCell(Sheet sheet, int columnIndex, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return null;
        return row.getCell(columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
    }
}
