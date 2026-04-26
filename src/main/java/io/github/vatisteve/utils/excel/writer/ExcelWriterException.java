package io.github.vatisteve.utils.excel.writer;

import java.io.IOException;

/**
 * Represents a custom exception specific to errors encountered during Excel file writing operations.
 * This exception extends {@link IOException}, allowing it to be used in scenarios
 * where input/output operations are involved in handling Excel files.
 * <p>
 * The primary purpose of this exception is to indicate and propagate meaningful
 * error messages whenever issues arise during the process of writing data to an Excel file.
 */
public class ExcelWriterException extends IOException {

    private static final long serialVersionUID = 3972340228891401352L;

    /**
     * Constructs a new {@code ExcelWriterException} with the specified detail message.
     *
     * @param message the detail message providing information about the exception.
     */
    public ExcelWriterException(String message) {
        super(message);
    }

}
