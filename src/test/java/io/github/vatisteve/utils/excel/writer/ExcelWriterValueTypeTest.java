package io.github.vatisteve.utils.excel.writer;

import io.github.vatisteve.utils.excel.ExcelTestSupport;
import io.github.vatisteve.utils.excel.ExcelUtilsFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.params.provider.Arguments.arguments;

/**
 * Exercises {@code ExcelWriterImpl}'s value-type dispatch (the {@code valueHandlers} map and
 * {@code detachAndSetCellValue}) for non-temporal types. Each case writes a single value and
 * verifies the resulting POI cell type and value. Temporal types live in
 * {@link ExcelWriterTemporalTypeTest} because they need a fixed time zone to be deterministic.
 */
@DisplayName("ExcelWriter value-type dispatch")
class ExcelWriterValueTypeTest {

    /** Writes one value into A1 of a fresh workbook and returns the reloaded cell for assertions. */
    private static void writeAndAssert(Object value, Consumer<Cell> cellAssertions) throws IOException {
        byte[] bytes;
        try (ExcelWriter writer = ExcelUtilsFactory.createExcelWriter()) {
            writer.startNewRow();
            writer.addCell(value);
            bytes = writer.build();
        }
        ExcelTestSupport.assertFirstSheet(bytes, sheet -> cellAssertions.accept(sheet.getRow(0).getCell(0)));
    }

    static Stream<Arguments> scalarValues() {
        return Stream.of(
                arguments("Boolean -> BOOLEAN", Boolean.TRUE,
                        (Consumer<Cell>) c -> {
                            assertEquals(CellType.BOOLEAN, c.getCellType());
                            assertTrue(c.getBooleanCellValue());
                        }),
                arguments("String -> STRING", "hello",
                        (Consumer<Cell>) c -> {
                            assertEquals(CellType.STRING, c.getCellType());
                            assertEquals("hello", c.getStringCellValue());
                        }),
                arguments("Byte -> NUMERIC", (byte) 5,
                        (Consumer<Cell>) c -> assertEquals(5d, c.getNumericCellValue())),
                arguments("Short -> NUMERIC", (short) 7,
                        (Consumer<Cell>) c -> assertEquals(7d, c.getNumericCellValue())),
                arguments("Integer -> NUMERIC", 42,
                        (Consumer<Cell>) c -> assertEquals(42d, c.getNumericCellValue())),
                arguments("Long -> NUMERIC", 123L,
                        (Consumer<Cell>) c -> assertEquals(123d, c.getNumericCellValue())),
                arguments("Float -> NUMERIC", 1.5f,
                        (Consumer<Cell>) c -> assertEquals(1.5d, c.getNumericCellValue())),
                arguments("Double -> NUMERIC", 2.5d,
                        (Consumer<Cell>) c -> assertEquals(2.5d, c.getNumericCellValue())),
                // Character has no dedicated Cell.setCellValue overload, so it widens to double.
                arguments("Character -> NUMERIC (widened to char code)", 'A',
                        (Consumer<Cell>) c -> assertEquals(65d, c.getNumericCellValue())),
                // BigDecimal/BigInteger are written as their plain string form to avoid precision loss.
                arguments("BigDecimal -> STRING (plain)", new BigDecimal("12345.6789"),
                        (Consumer<Cell>) c -> {
                            assertEquals(CellType.STRING, c.getCellType());
                            assertEquals("12345.6789", c.getStringCellValue());
                        }),
                arguments("BigInteger -> STRING", new BigInteger("99999999999999999999"),
                        (Consumer<Cell>) c -> {
                            assertEquals(CellType.STRING, c.getCellType());
                            assertEquals("99999999999999999999", c.getStringCellValue());
                        })
        );
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("scalarValues")
    void writesScalarValuesWithExpectedType(String name, Object value, Consumer<Cell> cellAssertions) throws IOException {
        writeAndAssert(value, cellAssertions);
    }

    @Test
    @DisplayName("null value produces a BLANK cell")
    void nullBecomesBlank() throws IOException {
        writeAndAssert(null, c -> assertEquals(CellType.BLANK, c.getCellType()));
    }

    @Test
    @DisplayName("Unhandled type falls back to toString()")
    void unhandledTypeFallsBackToToString() throws IOException {
        Object custom = new Object() {
            @Override
            public String toString() {
                return "CUSTOM-VALUE";
            }
        };
        writeAndAssert(custom, c -> {
            assertEquals(CellType.STRING, c.getCellType());
            assertEquals("CUSTOM-VALUE", c.getStringCellValue());
        });
    }
}
