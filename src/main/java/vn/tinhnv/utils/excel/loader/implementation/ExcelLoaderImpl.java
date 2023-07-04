package vn.tinhnv.utils.excel.loader.implementation;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;
import java.util.function.BinaryOperator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

import vn.tinhnv.utils.excel.helper.ExcelHelper;
import vn.tinhnv.utils.excel.loader.ExcelLoader;
import vn.tinhnv.utils.excel.loader.exception.CastCellValueExcelLoaderException;
import vn.tinhnv.utils.excel.loader.exception.ElementExcelLoaderNotFoundException;

public class ExcelLoaderImpl implements ExcelLoader {

    private Workbook workbook;
    private Sheet defaultSheet;

    private static final BinaryOperator<String> notFoundMessage = (t, s) -> String
            .format("There is no %s with '%s'", t, s);

    public ExcelLoaderImpl(InputStream iputStream) throws EncryptedDocumentException, IOException {
        workbook = WorkbookFactory.create(iputStream);
        defaultSheet = workbook.getSheetAt(0);
    }

    @Override
    public void setDefaultSheet(int s) throws ElementExcelLoaderNotFoundException {
        try {
            this.defaultSheet = workbook.getSheetAt(s);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        }
    }

    @Override
    public void setDefaultSheet(String s) throws ElementExcelLoaderNotFoundException {
        this.defaultSheet = Optional.ofNullable(workbook.getSheet(s))
                .orElseThrow(() -> new ElementExcelLoaderNotFoundException(notFoundMessage.apply("sheet name", s)));
    }

    @Override
    public Sheet getDefaultSheet() {
        return this.defaultSheet;
    }

    @Override
    public String getSheetName(int i) throws ElementExcelLoaderNotFoundException {
        try {
            return workbook.getSheetName(i);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        }
    }

    @Override
    public int getSheetIndex(String n) {
        return workbook.getSheetIndex(n);
    }

    @Override
    public long getLong(int s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).longValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public long getLong(String s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        Sheet workingSheet = workbook.getSheet(s);
        if (workingSheet == null)
            throw new ElementExcelLoaderNotFoundException(notFoundMessage.apply("sheet", s));

        try {
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).longValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public long getLong(int s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getLong(s, c.getColumn(), c.getRow());
    }

    @Override
    public long getLong(String s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getLong(s, c.getColumn(), c.getRow());
    }

    @Override
    public String getString(int s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return (String) ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public String getString(String s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        Sheet workingSheet = workbook.getSheet(s);
        if (workingSheet == null)
            throw new ElementExcelLoaderNotFoundException(notFoundMessage.apply("sheet", s));

        try {
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return (String) ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public String getString(int s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getString(s, c.getColumn(), c.getRow());
    }

    @Override
    public String getString(String s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getString(s, c.getColumn(), c.getRow());
    }

    @Override
    public int getInteger(int s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).intValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public int getInteger(String s, int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        Sheet workingSheet = workbook.getSheet(s);
        if (workingSheet == null)
            throw new ElementExcelLoaderNotFoundException(notFoundMessage.apply("sheet", s));

        try {
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).intValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public int getInteger(int s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(s, c.getColumn(), c.getRow());
    }

    @Override
    public int getInteger(String s, CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(s, c.getColumn(), c.getRow());
    }

    @Override
    public <T> T getValue(int s, int c, int r) throws ElementExcelLoaderNotFoundException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(String s, int c, int r) throws ElementExcelLoaderNotFoundException {
        Sheet workingSheet = workbook.getSheet(s);
        if (workingSheet == null)
            throw new ElementExcelLoaderNotFoundException(notFoundMessage.apply("sheet", s));

        try {
            Cell cell = ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(int s, CellAddress c) throws ElementExcelLoaderNotFoundException {
        return getValue(s, c.getColumn(), c.getRow());
    }

    @Override
    public <T> T getValue(String s, CellAddress c) throws ElementExcelLoaderNotFoundException {
        return getValue(s, c.getColumn(), c.getRow());
    }

    @Override
    public long getLong(int c, int r) throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Cell cell = ExcelHelper.getCell(defaultSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).longValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public long getLong(CellAddress c) throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getLong(c.getColumn(), c.getRow());
    }

    @Override
    public String getString(int c, int r)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Cell cell = ExcelHelper.getCell(defaultSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return (String) ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public String getString(CellAddress c)
            throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getString(c.getColumn(), c.getRow());
    }

    @Override
    public int getInteger(int c, int r) throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        try {
            Cell cell = ExcelHelper.getCell(defaultSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ((Double) ExcelHelper.getCellValue(cell)).intValue();
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public int getInteger(CellAddress c) throws ElementExcelLoaderNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(c.getColumn(), c.getRow());
    }

    @Override
    public <T> T getValue(int c, int r) throws ElementExcelLoaderNotFoundException {
        try {
            Cell cell = ExcelHelper.getCell(defaultSheet, c, r);
            if (cell == null)
                throw new ElementExcelLoaderNotFoundException(
                        notFoundMessage.apply("cell address", "[" + c + "," + r + "]"));
            return ExcelHelper.getCellValue(cell);
        } catch (IllegalArgumentException e) {
            throw new ElementExcelLoaderNotFoundException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(CellAddress c) throws ElementExcelLoaderNotFoundException {
        return getValue(c.getColumn(), c.getRow());
    }

    @Override
    public void close() throws IOException {
        this.workbook.close();
    }
}
