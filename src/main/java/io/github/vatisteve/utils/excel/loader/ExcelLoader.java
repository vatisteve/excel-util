package io.github.vatisteve.utils.excel.loader;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;

import java.io.Closeable;

/**
 * interface EcelLoader
 * <p>
 * Get the value from a specific cell in the working sheet
 * </p>
 *
 * @author Steve
 * @since May 23, 2023
 *
 */
public interface ExcelLoader extends Closeable {

    /**
     * Set the default working sheet with sheet index
     *
     * @param s the sheet index
     * @throws ElementNotFoundException if there is no sheet with this index {@code s}
     */
    void setDefaultSheet(int s) throws ElementNotFoundException;

    /**
     * Set the default working sheet with a sheet name
     *
     * @param s the sheet name
     * @throws ElementNotFoundException if there is no sheet with this name {@code s}
     */
    void setDefaultSheet(String s) throws ElementNotFoundException;

    /**
     * Get the default working sheet
     * @return {@code Sheet}
     */
    Sheet getDefaultSheet();

    /**
     * Get the sheet name by its index
     *
     * @param i the sheet index
     * @return the sheet name
     * @throws ElementNotFoundException if there is no sheet with this
     *                                  index
     */
    String getSheetName(int i) throws ElementNotFoundException;

    /**
     * Get the sheet index by its name
     *
     * @param n the sheet name
     * @return the sheet index. -1 if there is no sheet with {@code sheetName}
     */
    int getSheetIndex(String n);

    /**
     * Get a long value from a cell
     *
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a long value from a cell
     *
     * @param s the sheet name
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a long value from a cell
     *
     * @param s the sheet index
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a long value from a cell
     *
     * @param s the sheet name
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell
     *
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell
     *
     * @param s the sheet name
     * @param c the column index
     * @param r the row index
     * @return cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(String s, int c, int r)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell
     *
     * @param s the sheet index
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(int s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell
     *
     * @param s the sheet name
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(String s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell
     *
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell
     *
     * @param s the sheet name
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(String s, int c, int r)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell
     *
     * @param s the sheet index
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell
     *
     * @param s the sheet name
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(String s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a value from a cell
     *
     * @param <T> the specific type of value to get
     * @param s   the sheet index
     * @param c   the column index
     * @param r   the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(int s, int c, int r) throws ElementNotFoundException;

    /**
     * Get a value from a cell
     *
     * @param <T> the specific type of value to get
     * @param s   the sheet name
     * @param c   the column index
     * @param r   the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(String s, int c, int r) throws ElementNotFoundException;

    /**
     * Get a value from a cell
     *
     * @param <T> the specific type of value to get
     * @param s   the sheet index
     * @param c   the {@link CellAddress} from Apache POI
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(int s, CellAddress c) throws ElementNotFoundException;

    /**
     * Get a value from a cell
     *
     * @param <T> the specific type of value to get
     * @param s   the sheet name
     * @param c   the {@link CellAddress} from Apache POI
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(String s, CellAddress c) throws ElementNotFoundException;

    /**
     * Get a long value from a cell from the default sheet
     *
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a long value from a cell from the default sheet
     *
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Long}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Long}
     *                                           type or can't cast to {@code Long}
     */
    long getLong(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell from the default sheet
     *
     * @param c the column index
     * @param r the row index
     * @return cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a String value from a cell from the default sheet
     *
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type String
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell from the default sheet
     *
     * @param c the column index
     * @param r the row index
     * @return cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get an int value from a cell from the default sheet
     *
     * @param c the {@link CellAddress} from Apache POI
     * @return the cell value type {@code Integer}
     * @throws ElementNotFoundException          if there is no cell with that
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not {@code Integer}
     *                                           type or can't cast to {@code Integer}
     */
    int getInteger(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * Get a value from a cell from the default sheet
     *
     * @param <T> the specific type of value to get
     * @param c   the column index
     * @param r   the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(int c, int r) throws ElementNotFoundException;

    /**
     * Get a value from a cell from the default sheet
     *
     * @param <T> the specific type of value to get
     * @param c   the {@link CellAddress} from Apache POI
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                  address
     */
    <T> T getValue(CellAddress c) throws ElementNotFoundException;
}
