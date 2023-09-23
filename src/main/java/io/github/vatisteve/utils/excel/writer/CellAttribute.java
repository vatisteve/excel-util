package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

public final class CellAttribute {

    private final CellStyle cellStyle;
    private final Object value;

    private CellAttribute(Builder builder) {
        this.cellStyle = builder.cellStyle;
        this.value = builder.value;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public Object getValue() {
        return value;
    }

    public static final class Builder {
        private final CellStyle cellStyle;
        private Object value;
        public Builder(CellStyle style) {
            this.cellStyle = style;
        }
        public Builder value(Object value) {
            this.value = value;
            return this;
        }
        public CellAttribute build() {
            return new CellAttribute(this);
        }

    }
}
