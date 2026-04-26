package io.github.vatisteve.utils.excel.loader;

/**
 * Exception thrown when an error occurs during the casting of a cell value
 * in the context of Excel data loading operations.
 * <p>
 * This exception is used to indicate that a cell value in an Excel file could not
 * be appropriately cast to the desired data type or format required for processing.
 */
public final class CastCellValueExcelLoaderException extends Exception {

    private static final long serialVersionUID = 3192311250592917269L;

    /**
     * Constructor for CastCellValueExcelLoaderException.
     * @param message the detail message.
     */
    public CastCellValueExcelLoaderException(String message) {
        super(message);
    }
}
