package io.github.vatisteve.utils.excel.loader;

import java.io.Closeable;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;

/**
 * interface EcelLoader
 * <p>
 *  Get the value from a specific cell in the working sheet
 * </p>
 * @author Steve
 * @since May 23, 2023
 *
 */
public interface ExcelLoader extends Closeable {

    /**
     * @param s the sheet index
     * set the default working sheet with sheet index
     * @throws ElementNotFoundException if there is no sheet with this index {@code s}
     */
    void setDefaultSheet(int s) throws ElementNotFoundException;

    /**
     * @param s the sheet name
     * set the default working sheet with sheet name
     * @throws ElementNotFoundException if there is no sheet with this name {@code s}
     */
    void setDefaultSheet(String s) throws ElementNotFoundException;

    /**
     * @return the default working sheet {@code Sheet}
     */
    Sheet getDefaultSheet();

    /**
     * @param i the sheet index
     * @return the sheet name
     * @throws ElementNotFoundException if there is no sheet with this
     *                                             index
     */
    String getSheetName(int i) throws ElementNotFoundException;

    /**
     * @param n the sheet name
     * @return the sheet index. -1 if there is no sheet with {@code sheetName}
     */
    int getSheetIndex(String n);

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     */
    long getLong(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the rown index
     * @return cell valye type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     */
    long getLong(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     */
    long getLong(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     */
    long getLong(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     */
    String getString(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the rown index
     * @return cell valye type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     */
    String getString(String s, int c, int r)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     */
    String getString(int s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     */
    String getString(String s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     */
    int getInteger(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the rown index
     * @return cell valye type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     */
    int getInteger(String s, int c, int r)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     */
    int getInteger(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     */
    int getInteger(String s, CellAddress c)
            throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>   the specify type of value to get
     * @param s     the sheet index
     * @param c     the column index
     * @param r     the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     */
    <T> T getValue(int s, int c, int r) throws ElementNotFoundException;

    /**
     * @param <T>   the specify type of value to get
     * @param s     the sheet name
     * @param c     the column index
     * @param r     the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     */
    <T> T getValue(String s, int c, int r) throws ElementNotFoundException;

    /**
     * @param <T>   the specify type of value to get
     * @param s     the sheet index
     * @param c     the {@link CellAddress} from apache-poi
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     */
    <T> T getValue(int s, CellAddress c) throws ElementNotFoundException;

    /**
     * @param <T>   the specify type of value to get
     * @param s     the sheet name
     * @param c     the {@link CellAddress} from apache-poi
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     */
    <T> T getValue(String s, CellAddress c) throws ElementNotFoundException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    long getLong(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to long
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    long getLong(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    String getString(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to String
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    String getString(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    int getInteger(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * @throws CastCellValueExcelLoaderException   if the cell value is not long
     *                                             type or can't cast to int
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    int getInteger(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>   the specify type of value to get
     * @param c     the column index
     * @param r     the row index
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    <T> T getValue(int c, int r) throws ElementNotFoundException;

    /**
     * @param <T>   the specify type of value to get
     * @param c     the {@link CellAddress} from apache-poi
     * @return the value type {@code T}
     * @throws ElementNotFoundException if there is no cell with that
     *                                             address
     * <p> get cell value from sheet which has set before, if not the default
     *          sheet will be choose
     */
    <T> T getValue(CellAddress c) throws ElementNotFoundException;
}
