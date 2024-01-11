package io.github.vatisteve.utils.excel.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public final class ExcelHelper {

    private ExcelHelper() {
    }

    /**
     * @param cell the {@link Cell} to get value
     * @param tClass cell value's Java type
     * @return cell value, null if cell's type is {@link CellType#_NONE} or {@link CellType#ERROR}
     * @param <T> cell value
     */
    public static <T> T getCellValue(Cell cell, Class<T> tClass) {
        CellType cellType = cell.getCellTypeEnum();
        if (cellType == CellType._NONE) {
            // Cell does not exist
            return null;
        } else if (cellType == CellType.ERROR) {
            // Cell stores error value
            return null;
        } else if (cellType == CellType.FORMULA) {
            return getCellValue(cell, cell.getCachedFormulaResultTypeEnum(), tClass);
        } else {
            return getCellValue(cell, cellType, tClass);
        }
    }

    /**
     * @param cell the {@link Cell} to get value
     * @param cellType the {@link CellType} which can store value
     * @param tClass cell value's Java type
     * @return cell value base on {@code cellType} and {@code tClass}
     * @param <T> cell value
     */
    public static <T> T getCellValue(Cell cell, CellType cellType, Class<T> tClass) {
        switch (cellType) {
        case BOOLEAN:
            return ifStringCast(tClass, cell.getBooleanCellValue());
        case NUMERIC:
            return ifStringCast(tClass, cell.getNumericCellValue());
        case STRING:
            return ifStringCast(tClass, cell.getStringCellValue());
        case BLANK:
            return ifStringCast(tClass, null);
        default:
            return null;
        }
    }

    /**
     * Special casting for String value
     * @param tClass the Java type
     * @param o value to cast
     * @return value after casting
     * @param <T> returned value type
     */
    @SuppressWarnings("unchecked")
    private static <T> T ifStringCast(Class<T> tClass, Object o) {
        if (String.class.isAssignableFrom(tClass)) {
            return (T) (o != null ? String.valueOf(o) : "" /* empty string for string value */);
        }
        return o == null ? null : tClass.cast(o);
    }

    /**
     * @param sheet workbook sheet
     * @param columnIndex column index
     * @param rowIndex row index
     * @return null if there is no row with {@code rowIndex}, create new cell it does not exist
     */
    public static Cell getCell(Sheet sheet, int columnIndex, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return null;
        return row.getCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }

    /**
     * @param sheet workbook sheet
     * @param columnIndex column index
     * @param rowIndex row index
     * @see org.apache.poi.ss.util.SheetUtil#getCellWithMerges(Sheet, int, int)
     * @return null if there is no row with {@code rowIndex}, check existing cell in merged regions and create new cell it does not exist
     */
    public static Cell getCellWithMerges(Sheet sheet, int columnIndex, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            // trust blank as null cell to check in merge regions
            Cell cell = row.getCell(columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null) return cell;
        }
        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                row = sheet.getRow(mergedRegion.getFirstRow());
                if (row != null) {
                    return row.getCell(mergedRegion.getFirstColumn());
                }
            }
        }
        return getCell(sheet, columnIndex, rowIndex);
    }
}
