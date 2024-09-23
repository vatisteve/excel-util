package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.time.ZoneId;

public interface ExcelWriterConfiguration {

    /**
     * Sheet name
     */
    default String sheetName(int index) {
        return String.format("Data %d", index);
    }

    /**
     * Default format for local time instance
     */
    default String timeFormat() {
        return "HH:mm:ss";
    }

    /**
     * ZoneId default when convert ZoneDateTime to LocalDateTime
     */
    default ZoneId zoneId() {
        return ZoneId.systemDefault();
    }

    /**
     * Default cell style
     */
    default CellStyle cellStyle(Workbook activeWb) {
        return activeWb.createCellStyle();
    }

    /**
     * Excel Header
     */
    default ExcelHeader excelHeader(Workbook activeWb) {
        return null;
    }

    /**
     * Row height
     * @see org.apache.poi.ss.usermodel.Row#setHeight(short) 
     */
    default short rowHeight() {
        return -1;
    }

    final class DefaultConfiguration implements ExcelWriterConfiguration {}

    class ExcelHeader {

        private final String[] headers;
        private final CellStyle style;
        private final short height;
        private final int sheetIndex;

        private ExcelHeader(Builder builder) {
            this.headers = builder.headers;
            this.style = builder.style;
            this.height = builder.height;
            this.sheetIndex = builder.sheetIndex;
        }

        public String[] getHeaders() {
            return headers;
        }

        public CellStyle getStyle() {
            return style;
        }

        public short getHeight() {
            return height;
        }

        public int getSheetIndex() {
            return sheetIndex;
        }

        public static final class Builder {

            private String[] headers;
            private CellStyle style;
            private short height = -1;
            private int sheetIndex;

            public Builder headers(String... headers) {
                this.headers = headers;
                return this;
            }

            public Builder style(CellStyle style) {
                this.style = style;
                return this;
            }

            public Builder height(short height) {
                this.height = height;
                return this;
            }

            public Builder sheetIndex(int sheetIndex) {
                this.sheetIndex = sheetIndex;
                return this;
            }

            public ExcelHeader build() {
                return new ExcelHeader(this);
            }

        }
    }
}
