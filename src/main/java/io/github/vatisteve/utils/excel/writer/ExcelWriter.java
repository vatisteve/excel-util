package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

import java.io.Closeable;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Calendar;
import java.util.Date;

public interface ExcelWriter extends Closeable {
    void startAtSheet(int sheetIndex, int rowIndex, int columnIndex);
    void startNewRow();
    void startAtRow(int rowIndex);
    void addCell(CellAttribute cell);
    void setCellStyle(CellStyle style);
    void autoIncrementCell();
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
    // wrapper classes
    void addCell(Byte value);
    void addCell(Integer value);
    void addCell(Short value);
    void addCell(Long value);
    void addCell(Float value);
    void addCell(Double value);
    void addCell(Boolean value);
    void addCell(Character value);
    // Date and time
    void addCell(Instant value);
    void addCell(ZonedDateTime value);
    void addCell(OffsetDateTime value);
    void addCell(Date value);
    void addCell(LocalDate value);
    void addCell(LocalTime value);
    void addCell(LocalDateTime value);
    void addCell(Calendar value);
    // Other types
    void addCell(BigDecimal value);
    void addCell(BigInteger value);
}
