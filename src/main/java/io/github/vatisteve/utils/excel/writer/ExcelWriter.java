package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;

/**
 * An interface for creating and writing data to Excel files. This interface provides
 * methods for configuring sheets, rows, and cells within an Excel workbook. It
 * supports writing various data types to cells, applying styles, and exporting
 * the workbook to a byte array or an output stream.
 */
public interface ExcelWriter extends Closeable {

    /**
     * Retrieves the workbook instance.
     *
     * @return The Workbook object representing a spreadsheet or workbook.
     */
    Workbook getWorkbook();

    // initializer
    /**
     * Selects the active sheet and moves the cursor to the given row and column position.
     * <p>
     * This only sets the sheet and cursor indices; it does not make a row active. Call
     * {@link #startNewRow()} or {@link #startAtRow(int)} before adding cells, otherwise the
     * first {@code addCell(...)} will throw {@link IllegalStateException}.
     *
     * @param sheetIndex the index of the sheet within the workbook to move to (zero-based)
     * @param rowIndex the index of the row within the sheet to position at (zero-based)
     * @param columnIndex the index of the column within the sheet to position at (zero-based)
     * @throws ElementNotFoundException if the specified sheet cannot be found
     */
    void startAtSheet(int sheetIndex, int rowIndex, int columnIndex) throws ElementNotFoundException;

    /**
     * Creates a new row at the next row position, makes it the active row, and resets the
     * column cursor to the first column.
     */
    void startNewRow();

    /**
     * Creates a new row at the next row position with the given height, makes it the active
     * row, and resets the column cursor to the first column.
     * @param height the height of the row
     */
    void startNewRow(short height);

    /**
     * Positions the cursor on an existing row, makes it the active row, and resets the column
     * cursor to the first column.
     * <p>
     * Because the underlying workbook streams rows to disk, only rows still held in the
     * streaming window — i.e. rows created earlier in this writing session — can be revisited.
     * Rows that already existed in a template loaded from an {@code InputStream} are not
     * reachable and will be reported as not found.
     *
     * @param index the index of the row within the sheet to position at (zero-based)
     * @throws ElementNotFoundException if the specified row is not present in the streaming window
     */
    void startAtRow(int index) throws ElementNotFoundException;

    /**
     * Positions the cursor on an existing row, applies the given height, makes it the active
     * row, and resets the column cursor to the first column. The same streaming-window
     * limitation described on {@link #startAtRow(int)} applies.
     * @param index the index of the row within the sheet to position at (zero-based)
     * @param height the height of the row
     * @throws ElementNotFoundException if the specified row is not present in the streaming window
     */
    void startAtRow(int index, short height) throws ElementNotFoundException;

    // functions

    /**
     * Add a cell to the current row.
     * @param attribute the {@link CellAttribute} to add
     */
    void addCell(CellAttribute attribute);

    /**
     * Add a cell to the current row.
     * @param value the value to add
     */
    void addCell(Object value);

    /**
     * Add a cell to the current row.
     * @param value the value to add
     * @param style the {@link CellStyle} to set
     */
    void addCell(Object value, CellStyle style);

    /**
     * Set the cell style.
     * @param style the {@link CellStyle} to set
     */
    void setCellStyle(CellStyle style);

    /**
     * Auto increment the current cell.
     */
    void autoIncrementCell();

    /**
     * Auto increment the current cell.
     * @param style the {@link CellStyle} to set
     */
    void autoIncrementCell(CellStyle style);

    // build

    /**
     * Build the workbook and return the byte array.
     * @return the byte array containing the workbook data
     * @throws ExcelWriterException if an error occurs during the build process
     */
    byte[] build() throws ExcelWriterException;

    /**
     * Constructs and writes the content to the specified output stream.
     *
     * @param outputStream the output stream where the content will be written
     * @throws IOException if an I/O error occurs during writing
     */
    void build(OutputStream outputStream) throws IOException;
}
