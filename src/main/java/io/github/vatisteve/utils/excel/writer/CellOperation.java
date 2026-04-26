package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Represents a functional interface for applying custom operations to an Excel cell.
 * <p>
 * Implementations of this interface should provide logic to modify or process a specific cell
 * within a sheet. These operations may include setting values, applying styles, or performing
 * more advanced modifications based on the current state of the cell.
 * <p>
 * The {@code operate} method is the single abstract method (SAM) of this functional interface,
 * allowing it to be used in lambda expressions or method references.
 */
@FunctionalInterface
public interface CellOperation {

    /**
     * Applies a custom operation to a specific cell within a sheet.
     * The operation may involve modifying the cell's value, applying styles, or other transformations.
     *
     * @param sheet the sheet containing the cell to be operated on
     * @param cell  the cell to which the custom operation will be applied
     * @return the modified cell after the operation is applied
     * @throws ExcelWriterException if an error occurs during the operation
     */
    Cell operate(Sheet sheet, Cell cell) throws ExcelWriterException;

}
