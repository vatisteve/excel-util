package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.enumeration.ElementIdentifier;
import io.github.vatisteve.utils.excel.enumeration.ExcelElement;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;

public class ExcelWriterImpl implements ExcelWriter {

    private final Workbook workbook;
    private final ExcelWriterConfiguration configuration;
    private final CellStyle defaultCellStyle;
    private Sheet sheet = null;
    private Row currentRow = null;
    private int nextRowIdx = 0;
    private int nextColumnIdx = 0;
    private int cellIncrement = 1;

    public ExcelWriterImpl(InputStream is, ExcelWriterConfiguration configuration) throws ExcelWriterException {
        try {
            this.workbook = new SXSSFWorkbook(new XSSFWorkbook(is), -1, true);
        } catch (IOException e) {
            throw new ExcelWriterException(e.getMessage());
        }
        this.configuration = configuration;
        this.defaultCellStyle = this.configuration.defaultCellStyle(workbook);
        initHeader();
    }

    public ExcelWriterImpl(ExcelWriterConfiguration configuration) {
        this.workbook = new SXSSFWorkbook();
        this.configuration = configuration;
        this.sheet = workbook.createSheet();
        this.workbook.setSheetName(0, configuration.sheetName(0));
        this.defaultCellStyle = configuration.defaultCellStyle(workbook);
        initHeader();
    }

    private Cell switchToNewCell() {
        Cell cell = currentRow.getCell(nextColumnIdx++, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        if (defaultCellStyle != null) cell.setCellStyle(defaultCellStyle);
        return cell;
    }

    private Cell switchToNewCell(CellStyle style) {
        Cell newCell = currentRow.getCell(nextColumnIdx++, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        newCell.setCellStyle(style);
        return newCell;
    }

    private void initRowAttribute() {
        currentRow.setHeight(configuration.defaultRowHeight());
        nextRowIdx++;
        nextColumnIdx = 0;
    }

    private void initRowAttribute(short height) {
        currentRow.setHeight(height);
        nextRowIdx++;
        nextColumnIdx = 0;
    }

    private void initHeader() {
        ExcelWriterConfiguration.ExcelHeader header = configuration.excelHeader(workbook);
        if (header != null) {
            startNewRow(header.getHeight());
            CellStyle headerStyle = header.getStyle();
            if (headerStyle == null) headerStyle = defaultCellStyle;
            for (String h : header.getHeaders()) {
                switchToNewCell(headerStyle).setCellValue(h);
            }
        }
    }

    private void detachCellValue(Object value, CellStyle style) {
        Cell cell = switchToNewCell(style);
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Byte) {
            cell.setCellValue((Byte) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Short) {
            cell.setCellValue((Short) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Character) {
            cell.setCellValue((Character) value);
        } else if (value instanceof Instant) {
            cell.setCellValue(fromInstant((Instant) value));
        } else if (value instanceof ZonedDateTime) {
            cell.setCellValue(fromZonedDateTime((ZonedDateTime) value));
        } else if (value instanceof OffsetDateTime) {
            cell.setCellValue(fromOffsetDateTime((OffsetDateTime) value));
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof LocalTime) {
            cell.setCellValue(fromLocalTime((LocalTime) value));
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(fromBigDecimal((BigDecimal) value));
        } else if (value instanceof BigInteger) {
            cell.setCellValue(fromBigInteger((BigInteger) value));
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private LocalDateTime fromInstant(Instant value) {
        return fromZonedDateTime(value.atZone(configuration.defaultZoneId()));
    }

    private static LocalDateTime fromZonedDateTime(ZonedDateTime value) {
        return value.toLocalDateTime();
    }

    private static LocalDateTime fromOffsetDateTime(OffsetDateTime value) {
        return value.toLocalDateTime();
    }

    private String fromLocalTime(LocalTime value) {
        return DateTimeFormatter.ofPattern(configuration.defaultLocalTimeFormat()).format(value);
    }

    private static String fromBigDecimal(BigDecimal value) {
        return value.toPlainString();
    }

    private static String fromBigInteger(BigInteger value) {
        return value.toString();
    }

    @Override
    public Workbook getWorkbook() {
        return workbook;
    }

    @Override
    public void startAtSheet(int sheetIndex, int rowIndex, int columnIndex) throws ElementNotFoundException {
        try {
            this.sheet = workbook.getSheetAt(sheetIndex);
            nextRowIdx = rowIndex;
            nextColumnIdx = columnIndex;
        } catch (IllegalArgumentException e) {
            throw new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.POSITION, sheetIndex);
        }
    }

    @Override
    public void startNewRow() {
        currentRow = sheet.createRow(nextRowIdx);
        initRowAttribute();
    }

    @Override
    public void startAtRow(int index) throws ElementNotFoundException {
        currentRow = sheet.getRow(index);
        if (currentRow != null) {
            nextRowIdx = index;
            initRowAttribute();
        } else {
            throw new ElementNotFoundException(ExcelElement.ROW, ElementIdentifier.POSITION, index);
        }
    }

    @Override
    public void startNewRow(short height) {
        currentRow = sheet.createRow(nextRowIdx);
        initRowAttribute(height);
    }

    @Override
    public void startAtRow(int index, short height) throws ElementNotFoundException {
        currentRow = sheet.getRow(index);
        if (currentRow != null) {
            nextRowIdx = index;
            initRowAttribute(height);
        } else {
            throw new ElementNotFoundException(ExcelElement.ROW, ElementIdentifier.POSITION, index);
        }
    }

    @Override
    public void addCell(CellAttribute attribute) {
        if (attribute.getValue() != null) {
            detachCellValue(attribute.getValue(), attribute.getCellStyle());
        } else {
            switchToNewCell().setCellStyle(attribute.getCellStyle());
        }
    }

    @Override
    public void setCellStyle(CellStyle style) {
        Cell cell = currentRow.getCell(nextColumnIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setCellStyle(style);
    }

    @Override
    public void autoIncrementCell() {
        switchToNewCell().setCellValue(cellIncrement++);
    }

    @Override
    public void autoIncrementCell(CellStyle style) {
        switchToNewCell(style).setCellValue(cellIncrement++);
    }

    @Override
    public void addCell(String value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(byte value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(int value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(short value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(long value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(float value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(double value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(boolean value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(char value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(String value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(byte value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(int value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(short value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(long value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(float value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(double value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(boolean value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(char value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Byte value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Integer value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Short value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Long value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Float value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Double value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Boolean value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Character value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Byte value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Integer value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Short value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Long value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Float value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Double value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Boolean value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Character value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Instant value) {
        switchToNewCell().setCellValue(fromInstant(value));
    }

    @Override
    public void addCell(ZonedDateTime value) {
        switchToNewCell().setCellValue(fromZonedDateTime(value));
    }

    @Override
    public void addCell(OffsetDateTime value) {
        switchToNewCell().setCellValue(fromOffsetDateTime(value));
    }

    @Override
    public void addCell(Date value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(LocalDate value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(LocalTime value) {
        switchToNewCell().setCellValue(fromLocalTime(value));
    }

    @Override
    public void addCell(LocalDateTime value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Calendar value) {
        switchToNewCell().setCellValue(value);
    }

    @Override
    public void addCell(Instant value, CellStyle style) {
        switchToNewCell(style).setCellValue(fromInstant(value));
    }

    @Override
    public void addCell(ZonedDateTime value, CellStyle style) {
        switchToNewCell().setCellValue(fromZonedDateTime(value));
    }

    @Override
    public void addCell(OffsetDateTime value, CellStyle style) {
        switchToNewCell().setCellValue(fromOffsetDateTime(value));
    }

    @Override
    public void addCell(Date value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(LocalDate value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(LocalTime value, CellStyle style) {
        switchToNewCell(style).setCellValue(fromLocalTime(value));
    }

    @Override
    public void addCell(LocalDateTime value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(Calendar value, CellStyle style) {
        switchToNewCell(style).setCellValue(value);
    }

    @Override
    public void addCell(BigDecimal value) {
        switchToNewCell().setCellValue(fromBigDecimal(value));
    }

    @Override
    public void addCell(BigInteger value) {
        switchToNewCell().setCellValue(fromBigInteger(value));
    }

    @Override
    public void addCell(BigDecimal value, CellStyle style) {
        switchToNewCell(style).setCellValue(fromBigDecimal(value));
    }

    @Override
    public void addCell(BigInteger value, CellStyle style) {
        switchToNewCell(style).setCellValue(fromBigInteger(value));
    }

    @Override
    public byte[] build() throws ExcelWriterException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
            return bos.toByteArray();
        } catch (IOException e) {
            throw new ExcelWriterException(String.format("IOException occurred: %s", e.getMessage()));
        }
    }

    @Override
    public void build(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    @Override
    public void close() throws IOException {
        this.workbook.close();
    }
}
