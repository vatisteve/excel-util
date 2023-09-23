package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.time.ZoneId;

public interface ExcelWriterConfiguration {

    default String sheetName(int index) {
       return String.format("Data %d", index);
    }

    String defaultLocalTimeFormat();

    ZoneId defaultZoneId();

    CellStyle defaultCellStyle(Workbook activeWb);

    ExcelHeader excelHeader(Workbook activeWb);

    short defaultRowHeight();

    final class DefaultConfiguration implements ExcelWriterConfiguration {

        @Override
        public String defaultLocalTimeFormat() {
            return "HH:mm:ss";
        }

        @Override
        public ZoneId defaultZoneId() {
            return ZoneId.systemDefault();
        }

        @Override
        public CellStyle defaultCellStyle(Workbook activeWb) {
            return null;
        }

        @Override
        public ExcelHeader excelHeader(Workbook activeWb) {
            return null;
        }

        @Override
        public short defaultRowHeight() {
            return -1;
        }
    }

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
