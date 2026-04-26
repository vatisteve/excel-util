package io.github.vatisteve.utils.excel.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * A utility class that provides helper methods for working with Excel sheets
 * and cells using the Apache POI library. This class includes methods to
 * retrieve cell values and perform cell lookups within a sheet.
 *<p>
 * This class is final and cannot be instantiated.
 */
public final class ExcelHelper {

    private ExcelHelper() {
    }

    /**
     * Retrieves the value of a given cell from an Excel sheet. The value is returned as the specified type
     * based on the cell's type. If the cell contains a formula, the cached formula result type is used to retrieve the value.
     *
     * @param <T> The expected return type of the cell's value.
     * @param cell The cell from which the value is to be retrieved. Must not be null.
     * @return The value of the cell as an object of type {@code T}. Returns {@code null} if the cell type is
     *         {@code CellType._NONE} or {@code CellType.ERROR}.
     */
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

    /**
     * Retrieves the value of a given cell from an Excel sheet based on its specified {@code CellType}.
     * The method attempts to cast the cell value to the desired type {@code T}. If the cell type
     * is unsupported, the method returns {@code null}.
     *
     * @param <T> The expected return type of the cell's value.
     * @param cell The cell from which the value is to be retrieved. This should not be {@code null}.
     * @param cellType The specific type of the cell, used to determine how the value should be extracted.
     * @return The value of the cell as an object of type {@code T}. Returns {@code null} if the cell type is
     *         unsupported or not specified.
     */
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

    /**
     * Retrieves the cell at the specified column and row indices from the given Excel sheet.
     * If the row does not exist or the cell does not exist, this method returns {@code null}.
     *
     * @param sheet The Excel sheet from which the cell is to be retrieved. Must not be {@code null}.
     * @param columnIndex The zero-based index of the column where the cell is located.
     * @param rowIndex The zero-based index of the row where the cell is located.
     * @return The {@code Cell} object at the specified location in the sheet, or {@code null}
     *         if the row or cell does not exist.
     */
    public static Cell getCell(Sheet sheet, int columnIndex, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return null;
        return row.getCell(columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
    }

}
