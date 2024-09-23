package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

@FunctionalInterface
public interface CellOperation {
    Cell operate(Sheet sheet, Cell cell) throws ExcelWriterException;
}
