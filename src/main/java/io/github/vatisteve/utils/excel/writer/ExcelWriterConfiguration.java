package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.time.ZoneId;

/**
 * Interface providing configuration options for writing Excel files.
 * This interface defines default methods for customizing various
 * aspects of the Excel workbook, such as sheet names, cell styles,
 * header configuration, time format, and more.
 */
public interface ExcelWriterConfiguration {

    /**
     * Define the sheet name base on its index
     * @param index Sheet index
     * @return Sheet name
     */
    default String sheetName(int index) {
        return String.format("Data %d", index);
    }

    /**
     * Default format for local time instance
     * @return Time string format {@code HH:mm:ss}
     */
    default String timeFormat() {
        return "HH:mm:ss";
    }

    /**
     * ZoneId default when convert ZoneDateTime to LocalDateTime
     * @return the system default zone id
     */
    default ZoneId zoneId() {
        return ZoneId.systemDefault();
    }

    /**
     * Default cell style
     * @param activeWb  Active workbook to create cell style
     * @return a new instance cell style with no options
     */
    default CellStyle cellStyle(Workbook activeWb) {
        return activeWb.createCellStyle();
    }

    /**
     * Excel Header
     * @param activeWb  Active workbook to create excel header and its style
     * @return {@link ExcelHeader}
     */
    default ExcelHeader excelHeader(Workbook activeWb) {
        return null;
    }

    /**
     * Row height
     * @see org.apache.poi.ss.usermodel.Row#setHeight(short)
     * @return {@code -1} by default for row height
     */
    default short rowHeight() {
        return -1;
    }

    /**
     * A final implementation of the {@link ExcelWriterConfiguration} interface,
     * providing default configurations for writing Excel files.
     * <p>
     * This class offers a concrete base implementation of the configuration options
     * defined in the {@code ExcelWriterConfiguration} interface. It enables the use
     * of default settings for common aspects of Excel file generation, such as:
     * - Sheet naming conventions.
     * - Default time formatting.
     * - Zone ID for date and time conversions.
     * - Cell styles.
     * - Row height preferences.
     * - Excel header definitions.
     * <p>
     * The class is immutable and provides a straightforward way to leverage
     * default configurations without requiring additional customization.
     */
    final class DefaultConfiguration implements ExcelWriterConfiguration {
        /**
         * Constructs a {@code DefaultConfiguration} instance with predefined
         * default settings for Excel file generation.
         * <p>
         * This constructor initializes the {@code DefaultConfiguration} object,
         * providing a ready-to-use configuration state with default values
         * for various aspects of Excel writing. It requires no parameters,
         * ensuring a straightforward instantiation process.
         */
        public DefaultConfiguration() {
            // Default constructor
        }
    }

    /**
     * Represents a configuration for the header row in an Excel sheet. This class encapsulates
     * properties such as the header text, cell styling, row height, and the index of the sheet
     * to which the header belongs.
     * <p>
     * Instances of this class are immutable and can only be created using the {@link Builder}.
     */
    class ExcelHeader {

        private final String[] headers;
        private final CellStyle style;
        private final short height;
        private final int sheetIndex;

        /**
         * Constructs an instance of the ExcelHeader class using the provided Builder.
         *
         * @param builder the Builder object containing the configuration for the ExcelHeader,
         *                including header values, styling, row height, and sheet index.
         */
        private ExcelHeader(Builder builder) {
            this.headers = builder.headers;
            this.style = builder.style;
            this.height = builder.height;
            this.sheetIndex = builder.sheetIndex;
        }

        /**
         * Get the header values
         * @return the header values
         */
        public String[] getHeaders() {
            return headers;
        }

        /**
         * Get the cell style
         * @return the cell style
         */
        public CellStyle getStyle() {
            return style;
        }

        /**
         * Get the row height
         * @return the row height
         */
        public short getHeight() {
            return height;
        }

        /**
         * Get the sheet index
         * @return the sheet index
         */
        public int getSheetIndex() {
            return sheetIndex;
        }

        /**
         * A builder class for constructing immutable instances of {@link ExcelHeader}.
         */
        public static final class Builder {

            private String[] headers;
            private CellStyle style;
            private short height = -1;
            private int sheetIndex;

            /**
             * Default constructor for the Builder class.
             * Initializes a new Builder instance for constructing immutable instances of ExcelHeader.
             */
            public Builder() {
                // Default constructor
            }

            /**
             * Set the header values
             * @param headers the header values
             * @return the Builder instance for method chaining
             */
            public Builder headers(String... headers) {
                this.headers = headers;
                return this;
            }

            /**
             * Set the cell style
             * @param style the cell style
             * @return the Builder instance for method chaining
             */
            public Builder style(CellStyle style) {
                this.style = style;
                return this;
            }

            /**
             * Set the row height
             * @param height the row height
             * @return the Builder instance for method chaining
             */
            public Builder height(short height) {
                this.height = height;
                return this;
            }

            /**
             * Set the sheet index
             * @param sheetIndex the sheet index
             * @return the Builder instance for method chaining
             */
            public Builder sheetIndex(int sheetIndex) {
                this.sheetIndex = sheetIndex;
                return this;
            }

            /**
             * Build the ExcelHeader instance
             * @return the built ExcelHeader instance
             */
            public ExcelHeader build() {
                return new ExcelHeader(this);
            }

        }
    }
}
