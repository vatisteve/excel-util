package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

public final class CellAttribute {

    private final CellStyle cellStyle;
    private final Object value;
    private final CellOperation cellOperation;

    private CellAttribute(Builder builder) {
        this.cellStyle = builder.cellStyle;
        this.value = builder.value;
        this.cellOperation = builder.cellOperation;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public Object getValue() {
        return value;
    }

    public CellOperation getCellOperation() {
        return cellOperation;
    }

    public static final class Builder {
        private final CellStyle cellStyle;
        private Object value;
        private CellOperation cellOperation;
        public Builder(CellStyle style) {
            this.cellStyle = style;
        }
        public Builder() {
            this.cellStyle = null;
        }
        public Builder value(Object value) {
            this.value = value;
            return this;
        }
        public Builder cellOperation(CellOperation cellOperation) {
            this.cellOperation = cellOperation;
            return this;
        }
        public CellAttribute build() {
            return new CellAttribute(this);
        }

    }
}
