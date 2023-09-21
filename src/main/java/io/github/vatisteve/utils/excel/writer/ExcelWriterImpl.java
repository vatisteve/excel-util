package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Calendar;
import java.util.Date;

public class ExcelWriterImpl implements ExcelWriter {

    private final Workbook workbook;
    private final ExcelWriterConfiguration configuration;
    private short datetimeStyle;
    private short numberStyle;
    private Sheet sheet = null;
    private Row row = null;
    private int nextRowIdx = 0;
    private int nextColumnIdx = 0;
    private int cellIncrement = 1;

    public ExcelWriterImpl(InputStream is, ExcelWriterConfiguration configuration) throws IOException {
        this.workbook = new SXSSFWorkbook(new XSSFWorkbook(is), -1, true);
        this.configuration = configuration;
        initExcelStyleParams();
    }

    public ExcelWriterImpl(ExcelWriterConfiguration configuration) {
        this.workbook = new XSSFWorkbook();
        this.configuration = configuration;
        workbook.createSheet();
        initExcelStyleParams();
    }

    private void initExcelStyleParams() {
        DataFormat dataFormat = this.workbook.createDataFormat();
        this.datetimeStyle = dataFormat.getFormat(configuration.getDateTimeFormat());
        this.numberStyle = dataFormat.getFormat(configuration.getNumberFormat());
    }

    @Override
    public void startAtSheet(int sheetIndex, int rowIndex, int columnIndex) {

    }

    @Override
    public void startNewRow() {

    }

    @Override
    public void startAtRow(int rowIndex) {

    }

    @Override
    public void addCell(CellAttribute cell) {

    }

    @Override
    public void setCellStyle(CellStyle style) {

    }

    @Override
    public void autoIncrementCell() {

    }

    @Override
    public void addCell(String value) {

    }

    @Override
    public void addCell(byte value) {

    }

    @Override
    public void addCell(int value) {

    }

    @Override
    public void addCell(short value) {

    }

    @Override
    public void addCell(long value) {

    }

    @Override
    public void addCell(float value) {

    }

    @Override
    public void addCell(double value) {

    }

    @Override
    public void addCell(boolean value) {

    }

    @Override
    public void addCell(char value) {

    }

    @Override
    public void addCell(Byte value) {

    }

    @Override
    public void addCell(Integer value) {

    }

    @Override
    public void addCell(Short value) {

    }

    @Override
    public void addCell(Long value) {

    }

    @Override
    public void addCell(Float value) {

    }

    @Override
    public void addCell(Double value) {

    }

    @Override
    public void addCell(Boolean value) {

    }

    @Override
    public void addCell(Character value) {

    }

    @Override
    public void addCell(Instant value) {

    }

    @Override
    public void addCell(ZonedDateTime value) {

    }

    @Override
    public void addCell(OffsetDateTime value) {

    }

    @Override
    public void addCell(Date value) {

    }

    @Override
    public void addCell(LocalDate value) {

    }

    @Override
    public void addCell(LocalTime value) {

    }

    @Override
    public void addCell(LocalDateTime value) {

    }

    @Override
    public void addCell(Calendar value) {

    }

    @Override
    public void addCell(BigDecimal value) {

    }

    @Override
    public void addCell(BigInteger value) {

    }

    @Override
    public void close() throws IOException {

    }
}
