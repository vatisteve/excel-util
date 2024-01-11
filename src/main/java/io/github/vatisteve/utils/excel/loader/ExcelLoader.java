package io.github.vatisteve.utils.excel.loader;

import java.io.Closeable;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;

/**
 * Excel loader
 * <p>Get the value from a specific cell in the working sheet</p>
 *
 * @author Steve
 * @since 1.0.0
 */
public interface ExcelLoader extends Closeable {

    void loadCellValueWithTakingAccountOfMergedRegion(boolean flag);

    /**
     * @param s the sheet index
     *          set the default working sheet with sheet index
     * @throws ElementNotFoundException if there is no sheet with this index {@code s}
     */
    void setDefaultSheet(int s) throws ElementNotFoundException;

    /**
     * @param s the sheet name
     *          set the default working sheet with sheet name
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
     * @throws ElementNotFoundException if there is no sheet with given
     *                                  index
     */
    String getSheetName(int i) throws ElementNotFoundException;

    /**
     * @param n the sheet name
     * @return the sheet index, -1 if there is no sheet with {@code sheetName}
     */
    int getSheetIndex(String n);

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     */
    Long getLong(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the row index
     * @return cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     */
    Long getLong(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     */
    Long getLong(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     */
    Long getLong(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the rown index
     * @return cell valye type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with give
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     */
    String getString(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the column index
     * @param r the row index
     * @return cell value type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     */
    Integer getInteger(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the column index
     * @param r the rown index
     * @return cell valye type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     */
    Integer getInteger(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet index
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     */
    Integer getInteger(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param s the sheet name
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     */
    Integer getInteger(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param s      the sheet index
     * @param c      the column index
     * @param r      the row index
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell can't cast to {@code tClass}
     */
    <T> T getValue(int s, int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param s      the sheet name
     * @param c      the column index
     * @param r      the row index
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell can't cast to {@code tClass}
     */
    <T> T getValue(String s, int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param s      the sheet index
     * @param c      the {@link CellAddress} from apache-poi
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell can't cast to {@code tClass}
     */
    <T> T getValue(int s, CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param s      the sheet name
     * @param c      the {@link CellAddress} from apache-poi
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell can't cast to {@code tClass}
     */
    <T> T getValue(String s, CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    Long getLong(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type long, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Long
     *                                           type or can't cast to Long
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    Long getLong(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    String getString(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type String, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not String
     *                                           type or can't cast to String
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    String getString(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the column index
     * @param r the row index
     * @return cell value type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    Integer getInteger(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param c the {@link CellAddress} from apache-poi
     * @return the cell value type int, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value is not Integer
     *                                           type or can't cast to Integer
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    Integer getInteger(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param c      the column index
     * @param r      the row index
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value can't cast to {@code tClass}
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    <T> T getValue(int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;

    /**
     * @param <T>    the specify type of value to get
     * @param c      the {@link CellAddress} from apache-poi
     * @param tClass the class to cast cell value
     * @return the value type {@code T}, null if cell is blank
     * @throws ElementNotFoundException          if there is no cell with given
     *                                           address
     * @throws CastCellValueExcelLoaderException if the cell value can't cast to {@code tClass}
     *                                           <p> get cell value from sheet which has set before, if not the default
     *                                           sheet will be chosen
     */
    <T> T getValue(CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException;
}
