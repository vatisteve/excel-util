package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ElementNotFoundException;
import io.github.vatisteve.utils.excel.common.ElementIdentifier;
import io.github.vatisteve.utils.excel.common.ExcelElement;
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
import java.util.*;
import java.util.function.BiConsumer;

/**
 * Implementation of the ExcelWriter for writing data to Excel spreadsheets.
 * This class provides methods to write data to a workbook, manage rows and cells,
 * apply styles, and build the resulting Excel file.
 * <p>
 * The class supports various data types and allows customization of cell styles.
 */
public class ExcelWriterImpl implements ExcelWriter {

    private final Workbook workbook;
    private final ExcelWriterConfiguration configuration;
    private final CellStyle defaultCellStyle;
    private final DateTimeFormatter timeFormatter;
    private final Map<Class<?>, BiConsumer<Object, Cell>> valueHandlers;
    private Sheet sheet = null;
    private Row currentRow = null;
    private int nextRowIdx = 0;
    private int nextColumnIdx = 0;
    private int cellIncrement = 1;

    /**
     * Constructs an instance of ExcelWriterImpl, initializing the workbook and configuration.
     *
     * @param is the InputStream containing the Excel template to initialize the workbook
     * @param configuration the ExcelWriterConfiguration providing customizations for the Excel writer
     * @throws ExcelWriterException if an error occurs while creating or configuring the workbook
     */
    public ExcelWriterImpl(InputStream is, ExcelWriterConfiguration configuration) throws ExcelWriterException {
        try {
            this.workbook = new SXSSFWorkbook(new XSSFWorkbook(is), -1, true);
        } catch (IOException e) {
            throw new ExcelWriterException(e.getMessage());
        }
        this.configuration = configuration;
        this.defaultCellStyle = configuration.cellStyle(workbook);
        this.timeFormatter = DateTimeFormatter.ofPattern(configuration.timeFormat());
        this.valueHandlers = initValueHandlers();
        initHeader();
    }

    /**
     * Constructs an instance of ExcelWriterImpl, initializing the workbook, sheet, and default cell style
     * based on the provided configuration.
     *
     * @param configuration the ExcelWriterConfiguration instance that provides the initial setup
     *                       and customizations for the Excel writer, such as sheet naming
     *                       and default cell style.
     */
    public ExcelWriterImpl(ExcelWriterConfiguration configuration) {
        this.workbook = new SXSSFWorkbook();
        this.configuration = configuration;
        this.sheet = workbook.createSheet();
        this.workbook.setSheetName(0, configuration.sheetName(0));
        this.defaultCellStyle = configuration.cellStyle(workbook);
        this.timeFormatter = DateTimeFormatter.ofPattern(configuration.timeFormat());
        this.valueHandlers = initValueHandlers();
        initHeader();
    }

    private Map<Class<?>, BiConsumer<Object, Cell>> initValueHandlers() {
        Map<Class<?>, BiConsumer<Object, Cell>> handlers = new HashMap<>();
        handlers.put(Boolean.class, (v, c) -> c.setCellValue((Boolean) v));
        handlers.put(String.class, (v, c) -> c.setCellValue((String) v));
        handlers.put(Byte.class, (v, c) -> c.setCellValue((Byte) v));
        handlers.put(Integer.class, (v, c) -> c.setCellValue((Integer) v));
        handlers.put(Short.class, (v, c) -> c.setCellValue((Short) v));
        handlers.put(Long.class, (v, c) -> c.setCellValue((Long) v));
        handlers.put(Float.class, (v, c) -> c.setCellValue((Float) v));
        handlers.put(Double.class, (v, c) -> c.setCellValue((Double) v));
        handlers.put(Character.class, (v, c) -> c.setCellValue((Character) v));
        handlers.put(Instant.class, (v, c) -> c.setCellValue(fromInstant((Instant) v)));
        handlers.put(ZonedDateTime.class, (v, c) -> c.setCellValue(fromZonedDateTime((ZonedDateTime) v)));
        handlers.put(OffsetDateTime.class, (v, c) -> c.setCellValue(fromOffsetDateTime((OffsetDateTime) v)));
        handlers.put(Date.class, (v, c) -> c.setCellValue((Date) v));
        handlers.put(LocalDate.class, (v, c) -> c.setCellValue((LocalDate) v));
        handlers.put(LocalDateTime.class, (v, c) -> c.setCellValue((LocalDateTime) v));
        handlers.put(LocalTime.class, (v, c) -> c.setCellValue(fromLocalTime((LocalTime) v)));
        handlers.put(Calendar.class, (v, c) -> c.setCellValue((Calendar) v));
        handlers.put(BigDecimal.class, (v, c) -> c.setCellValue(fromBigDecimal((BigDecimal) v)));
        handlers.put(BigInteger.class, (v, c) -> c.setCellValue(fromBigInteger((BigInteger) v)));
        return handlers;
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
        currentRow.setHeight(configuration.rowHeight());
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

    private void detachAndSetCellValue(Object value, Cell cell) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        BiConsumer<Object, Cell> handler = valueHandlers.get(value.getClass());
        if (handler != null) {
            handler.accept(value, cell);
        } else {
            // Check for potential subclasses of handled types (e.g., java.sql.Date)
            valueHandlers.entrySet().stream()
                    .filter(entry -> entry.getKey().isInstance(value))
                    .findFirst()
                    .map(Map.Entry::getValue)
                    .orElseGet(() -> (v, c) -> c.setCellValue(v.toString()))
                    .accept(value, cell);
        }
    }

    private LocalDateTime fromInstant(Instant value) {
        return fromZonedDateTime(value.atZone(configuration.zoneId()));
    }

    private static LocalDateTime fromZonedDateTime(ZonedDateTime value) {
        return value.toLocalDateTime();
    }

    private static LocalDateTime fromOffsetDateTime(OffsetDateTime value) {
        return value.toLocalDateTime();
    }

    private String fromLocalTime(LocalTime value) {
        return timeFormatter.format(value);
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
        if (attribute.getCellOperation() != null) {
            try {
                Cell cell = attribute.getCellOperation().operate(sheet, switchToNewCell());
                setCellValue(attribute, cell);
                return; // complete job
            } catch (ExcelWriterException e) {
                // Add warning ...
                // Ignore exception and continue to the next job
            }
        }
        Cell cell = newCellFrom(attribute);
        setCellValue(attribute, cell);
    }

    private void setCellValue(CellAttribute attribute, Cell cell) {
        if (attribute.getValue() != null) {
            detachAndSetCellValue(attribute.getValue(), cell);
        }
    }

    private Cell newCellFrom(CellAttribute attribute) {
        return Optional.ofNullable(attribute.getCellStyle())
                .map(this::switchToNewCell).orElseGet(this::switchToNewCell);
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
    public void addCell(Object value) {
        detachAndSetCellValue(value, switchToNewCell());
    }

    @Override
    public void addCell(Object value, CellStyle style) {
        detachAndSetCellValue(value, switchToNewCell(style));
    }

    @Override
    public byte[] build() throws ExcelWriterException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
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
