package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

import java.time.ZoneId;

public interface ExcelWriterConfiguration {
    String defaultLocalTimeFormat();

    ZoneId defaultZoneId();

    CellStyle defaultCellStyle();

    ExcelHeader excelHeader();

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
        public CellStyle defaultCellStyle() {
            return null;
        }

        @Override
        public ExcelHeader excelHeader() {
            return null;
        }

        @Override
        public short defaultRowHeight() {
            return -1; // auto size
        }
    }

    class ExcelHeader {
        private final String[] headers;
        private final CellStyle style;
        private final short height;

        private ExcelHeader(Builder builder) {
            this.headers = builder.headers;
            this.style = builder.style;
            this.height = builder.height;
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

        public static final class Builder {
            private String[] headers;
            private CellStyle style;
            private short height = -1;

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

            public ExcelHeader build() {
                return new ExcelHeader(this);
            }
        }
    }
}
