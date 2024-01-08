package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Calendar;
import java.util.Date;

public interface ExcelWriter extends Closeable {
    Workbook getWorkbook();
    // initializer
    void startAtSheet(int sheetIndex, int rowIndex, int columnIndex) throws ElementNotFoundException;
    void startNewRow();
    void startNewRow(short height);
    void startAtRow(int index) throws ElementNotFoundException;
    void startAtRow(int index, short height) throws ElementNotFoundException;

    // functions
    void addCell(CellAttribute attribute);
    void setCellStyle(CellStyle style);
    void autoIncrementCell();
    void autoIncrementCell(CellStyle style);
    // primitive types
    void addCell(String value);
    void addCell(byte value);
    void addCell(int value);
    void addCell(short value);
    void addCell(long value);
    void addCell(float value);
    void addCell(double value);
    void addCell(boolean value);
    void addCell(char value);
    void addCell(String value, CellStyle style);
    void addCell(byte value, CellStyle style);
    void addCell(int value, CellStyle style);
    void addCell(short value, CellStyle style);
    void addCell(long value, CellStyle style);
    void addCell(float value, CellStyle style);
    void addCell(double value, CellStyle style);
    void addCell(boolean value, CellStyle style);
    void addCell(char value, CellStyle style);
    // wrapper classes
    void addCell(Byte value);
    void addCell(Integer value);
    void addCell(Short value);
    void addCell(Long value);
    void addCell(Float value);
    void addCell(Double value);
    void addCell(Boolean value);
    void addCell(Character value);
    void addCell(Byte value, CellStyle style);
    void addCell(Integer value, CellStyle style);
    void addCell(Short value, CellStyle style);
    void addCell(Long value, CellStyle style);
    void addCell(Float value, CellStyle style);
    void addCell(Double value, CellStyle style);
    void addCell(Boolean value, CellStyle style);
    void addCell(Character value, CellStyle style);
    // Date and time
    void addCell(Instant value);
    void addCell(ZonedDateTime value);
    void addCell(OffsetDateTime value);
    void addCell(Date value);
    void addCell(LocalTime value);
    void addCell(Calendar value);
    void addCell(Instant value, CellStyle style);
    void addCell(ZonedDateTime value, CellStyle style);
    void addCell(OffsetDateTime value, CellStyle style);
    void addCell(Date value, CellStyle style);
    void addCell(LocalTime value, CellStyle style);
    void addCell(Calendar value, CellStyle style);
    // Other types
    void addCell(BigDecimal value);
    void addCell(BigInteger value);
    void addCell(BigDecimal value, CellStyle style);
    void addCell(BigInteger value, CellStyle style);
    // build
    byte[] build() throws ExcelWriterException;

    void build(OutputStream outputStream) throws IOException;
}
