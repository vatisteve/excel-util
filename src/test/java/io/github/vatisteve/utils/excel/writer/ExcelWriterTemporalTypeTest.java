package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ExcelTestSupport;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Verifies how {@code ExcelWriterImpl} converts the date/time families. A fixed UTC time zone
 * is configured so {@link Instant} conversion (which depends on {@code configuration.zoneId()})
 * is deterministic. Excel stores dates as a floating-point serial, so comparisons are made at
 * second precision.
 */
@DisplayName("ExcelWriter temporal-type dispatch")
class ExcelWriterTemporalTypeTest {

    /** Config pinning the zone/format so date-time conversions are reproducible across machines. */
    private static final class UtcConfig implements ExcelWriterConfiguration {
        @Override
        public ZoneId zoneId() {
            return ZoneId.of("UTC");
        }

        @Override
        public String timeFormat() {
            return "HH:mm:ss";
        }
    }

    private static void writeAndAssert(Object value, Consumer<Cell> cellAssertions) throws IOException {
        byte[] bytes;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter(new UtcConfig())) {
            writer.startNewRow();
            writer.addCell(value);
            bytes = writer.build();
        }
        ExcelTestSupport.assertFirstSheet(bytes, sheet -> cellAssertions.accept(sheet.getRow(0).getCell(0)));
    }

    @Test
    @DisplayName("LocalDate keeps its calendar date")
    void localDate() throws IOException {
        LocalDate value = LocalDate.of(2023, 4, 5);
        writeAndAssert(value, c -> assertEquals(value, c.getLocalDateTimeCellValue().toLocalDate()));
    }

    @Test
    @DisplayName("LocalDateTime round-trips to the second")
    void localDateTime() throws IOException {
        LocalDateTime value = LocalDateTime.of(2023, 4, 5, 6, 7, 8);
        writeAndAssert(value, c ->
                assertEquals(value, c.getLocalDateTimeCellValue().truncatedTo(ChronoUnit.SECONDS)));
    }

    @Test
    @DisplayName("Instant is resolved against the configured zone (UTC)")
    void instant() throws IOException {
        Instant value = Instant.parse("2020-01-15T10:20:30Z");
        LocalDateTime expected = LocalDateTime.of(2020, 1, 15, 10, 20, 30);
        writeAndAssert(value, c ->
                assertEquals(expected, c.getLocalDateTimeCellValue().truncatedTo(ChronoUnit.SECONDS)));
    }

    @Test
    @DisplayName("ZonedDateTime is stored as its local date-time component")
    void zonedDateTime() throws IOException {
        ZonedDateTime value = ZonedDateTime.of(2021, 2, 3, 4, 5, 6, 0, ZoneId.of("America/New_York"));
        writeAndAssert(value, c ->
                assertEquals(value.toLocalDateTime(), c.getLocalDateTimeCellValue().truncatedTo(ChronoUnit.SECONDS)));
    }

    @Test
    @DisplayName("OffsetDateTime is stored as its local date-time component")
    void offsetDateTime() throws IOException {
        OffsetDateTime value = OffsetDateTime.of(2022, 3, 4, 5, 6, 7, 0, ZoneOffset.ofHours(2));
        writeAndAssert(value, c ->
                assertEquals(value.toLocalDateTime(), c.getLocalDateTimeCellValue().truncatedTo(ChronoUnit.SECONDS)));
    }

    @Test
    @DisplayName("java.util.Date round-trips to the second")
    void utilDate() throws IOException {
        Date value = Date.from(Instant.parse("2019-06-07T08:09:10Z"));
        writeAndAssert(value, c ->
                assertEquals(value.getTime() / 1000, c.getDateCellValue().getTime() / 1000));
    }

    @Test
    @DisplayName("Calendar round-trips its wall-clock value")
    void calendar() throws IOException {
        // Excel stores a zone-less wall-clock serial, so compare the wall-clock fields rather
        // than epoch millis (which would differ by the JVM's default-zone offset on read-back).
        Calendar value = new GregorianCalendar();
        value.clear();
        value.set(2024, Calendar.MARCH, 9, 11, 12, 13);
        writeAndAssert(value, c ->
                assertEquals(LocalDateTime.of(2024, 3, 9, 11, 12, 13),
                        c.getLocalDateTimeCellValue().truncatedTo(ChronoUnit.SECONDS)));
    }

    @Test
    @DisplayName("LocalTime is formatted as text using timeFormat()")
    void localTime() throws IOException {
        writeAndAssert(LocalTime.of(13, 45, 30), c -> {
            assertEquals(CellType.STRING, c.getCellType());
            assertEquals("13:45:30", c.getStringCellValue());
        });
    }

    @Test
    @DisplayName("java.sql.Date is routed to the Date handler via the subclass fallback")
    void sqlDateUsesSubclassFallback() throws IOException {
        java.sql.Date value = new java.sql.Date(Instant.parse("2018-01-02T00:00:00Z").toEpochMilli());
        // The discriminating check: a sql.Date is NOT an exact key in the handler map, so this only
        // lands on NUMERIC (a date serial) if the isInstance fallback matched java.util.Date.
        // A miss would fall through to toString() and produce a STRING cell instead.
        writeAndAssert(value, c -> assertEquals(CellType.NUMERIC, c.getCellType()));
    }
}
