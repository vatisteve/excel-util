package io.github.vatisteve.utils.excel.loader;

import io.github.vatisteve.utils.excel.enumeration.ElementIdentifier;
import io.github.vatisteve.utils.excel.enumeration.ExcelElement;
import io.github.vatisteve.utils.excel.helper.ExcelHelper;
import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;
import java.util.function.Function;

public class ExcelLoaderImpl implements ExcelLoader {

    private final Workbook workbook;
    private Sheet defaultSheet;
    private boolean loadCellValueWithTakingAccountOfMergedRegion;

    public ExcelLoaderImpl(InputStream inputStream) throws EncryptedDocumentException, IOException {
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (InvalidFormatException e) {
            throw new IOException(e);
        }
        defaultSheet = workbook.getSheetAt(0);
    }

    @Override
    public void loadCellValueWithTakingAccountOfMergedRegion(boolean flag) {
        this.loadCellValueWithTakingAccountOfMergedRegion = flag;
    }

    @Override
    public void setDefaultSheet(int s) throws ElementNotFoundException {
        try {
            this.defaultSheet = workbook.getSheetAt(s);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, s);
        }
    }

    @Override
    public void setDefaultSheet(String s) throws ElementNotFoundException {
        this.defaultSheet = Optional.ofNullable(workbook.getSheet(s)).orElseThrow(() -> new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, s));
    }

    @Override
    public Sheet getDefaultSheet() {
        return this.defaultSheet;
    }

    @Override
    public String getSheetName(int i) throws ElementNotFoundException {
        try {
            return workbook.getSheetName(i);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, i);
        }
    }

    @Override
    public int getSheetIndex(String n) {
        return workbook.getSheetIndex(n);
    }

    private <T> T castToNumber(Sheet sheet, int c, int r, Function<Double, T> transformFromDouble) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Cell cell = loadCellValueWithTakingAccountOfMergedRegion ? ExcelHelper.getCellWithMerges(sheet, c, r) : ExcelHelper.getCell(sheet, c, r);
            if (cell == null) throw new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, c, r);
            return Optional.ofNullable(ExcelHelper.getCellValue(cell, Number.class)).map(Number::doubleValue).map(transformFromDouble).orElse(null);
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    private String castToString(Sheet sheet, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Cell cell = loadCellValueWithTakingAccountOfMergedRegion ? ExcelHelper.getCellWithMerges(sheet, c, r) : ExcelHelper.getCell(sheet, c, r);
            if (cell == null) throw new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, c, r);
            return ExcelHelper.getCellValue(cell, String.class);
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    private <T> T castToNumber(int s, int c, int r, Function<Double, T> transformFromDouble) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            return castToNumber(workingSheet, c, r, transformFromDouble);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, s);
        }
    }

    private <T> T castToNumber(String s, int c, int r, Function<Double, T> transformFromDouble) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheet(s);
            return castToNumber(workingSheet, c, r, transformFromDouble);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.NAME, s);
        }
    }

    private String castToString(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            return castToString(workingSheet, c, r);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, s);
        }
    }

    private String castToString(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheet(s);
            return castToString(workingSheet, c, r);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.NAME, s);
        }
    }

    @Override
    public Long getLong(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(s, c, r, Double::longValue);
    }

    @Override
    public Long getLong(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(s, c, r, Double::longValue);
    }

    @Override
    public Long getLong(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getLong(s, c.getColumn(), c.getRow());
    }

    @Override
    public Long getLong(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getLong(s, c.getColumn(), c.getRow());
    }

    @Override
    public String getString(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToString(s, c, r);
    }

    @Override
    public String getString(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToString(s, c, r);
    }

    @Override
    public String getString(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getString(s, c.getColumn(), c.getRow());
    }

    @Override
    public String getString(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getString(s, c.getColumn(), c.getRow());
    }

    @Override
    public Integer getInteger(int s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(s, c, r, Double::intValue);
    }

    @Override
    public Integer getInteger(String s, int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(s, c, r, Double::intValue);
    }

    @Override
    public Integer getInteger(int s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(s, c.getColumn(), c.getRow());
    }

    @Override
    public Integer getInteger(String s, CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(s, c.getColumn(), c.getRow());
    }

    @Override
    public <T> T getValue(int s, int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheetAt(s);
            Cell cell = loadCellValueWithTakingAccountOfMergedRegion ? ExcelHelper.getCellWithMerges(workingSheet, c, r) : ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null) throw new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, c, r);
            return ExcelHelper.getCellValue(cell, tClass);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, s);
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(String s, int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        try {
            Sheet workingSheet = workbook.getSheet(s);
            Cell cell = loadCellValueWithTakingAccountOfMergedRegion ? ExcelHelper.getCellWithMerges(workingSheet, c, r) : ExcelHelper.getCell(workingSheet, c, r);
            if (cell == null) throw new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, c, r);
            return ExcelHelper.getCellValue(cell, tClass);
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.NAME, s);
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(int s, CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getValue(s, c.getColumn(), c.getRow(), tClass);
    }

    @Override
    public <T> T getValue(String s, CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getValue(s, c.getColumn(), c.getRow(), tClass);
    }

    @Override
    public Long getLong(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(defaultSheet, c, r, Double::longValue);
    }

    @Override
    public Long getLong(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getLong(c.getColumn(), c.getRow());
    }

    @Override
    public String getString(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToString(defaultSheet, c, r);
    }

    @Override
    public String getString(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getString(c.getColumn(), c.getRow());
    }

    @Override
    public Integer getInteger(int c, int r) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return castToNumber(defaultSheet, c, r, Double::intValue);
    }

    @Override
    public Integer getInteger(CellAddress c) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getInteger(c.getColumn(), c.getRow());
    }

    @Override
    public <T> T getValue(int c, int r, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        Cell cell = loadCellValueWithTakingAccountOfMergedRegion ? ExcelHelper.getCellWithMerges(defaultSheet, c, r) : ExcelHelper.getCell(defaultSheet, c, r);
        if (cell == null) throw new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, c, r);
        try {
            return ExcelHelper.getCellValue(cell, tClass);
        } catch (ClassCastException e) {
            throw new CastCellValueExcelLoaderException(e.getMessage());
        }
    }

    @Override
    public <T> T getValue(CellAddress c, Class<T> tClass) throws ElementNotFoundException, CastCellValueExcelLoaderException {
        return getValue(c.getColumn(), c.getRow(), tClass);
    }

    @Override
    public void close() throws IOException {
        this.workbook.close();
    }
}
