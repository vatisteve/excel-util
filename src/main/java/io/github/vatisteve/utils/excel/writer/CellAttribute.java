package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Represents the attributes of a cell in an Excel sheet, including its style, value, and any custom operation that should
 * be performed on the cell during write operations.
 *<p>
 * Instances of this class are immutable and can only be created using the {@link Builder}.
 *<p>
 * The {@link CellStyle} defines the formatting of the cell, the {@link Object} value contains the content,
 * and the {@link CellOperation} provides a functional interface for applying custom operations to the cell.
 */
public final class CellAttribute {

    private final CellStyle cellStyle;
    private final Object value;
    private final CellOperation cellOperation;

    private CellAttribute(Builder builder) {
        this.cellStyle = builder.cellStyle;
        this.value = builder.value;
        this.cellOperation = builder.cellOperation;
    }

    /**
     * Get the cell style.
     * @return the {@link CellStyle} associated with the cell.
     */
    public CellStyle getCellStyle() {
        return cellStyle;
    }

    /**
     * Get the cell value.
     * @return the value of the cell.
     */
    public Object getValue() {
        return value;
    }

    /**
     * Get the cell operation.
     * @return the {@link CellOperation} associated with the cell.
     */
    public CellOperation getCellOperation() {
        return cellOperation;
    }

    /**
     * A builder class for constructing immutable instances of {@link CellAttribute}.
     *<p>
     * The builder provides methods to set the attributes of a {@link CellAttribute}, including its style, value, and any custom
     * operation to be performed on the cell. It ensures controlled and flexible creation of {@link CellAttribute} objects.
     */
    public static final class Builder {
        private final CellStyle cellStyle;
        private Object value;
        private CellOperation cellOperation;

        /**
         * Constructs a Builder instance with the specified {@link CellStyle}.
         * @param style the {@link CellStyle} to be associated with the builder. This defines the stylistic
         *              attributes for the cells that will be constructed using this builder.
         */
        public Builder(CellStyle style) {
            this.cellStyle = style;
        }

        /**
         * Constructs a default {@code Builder} instance with no associated {@link CellStyle}.
         * <p>
         * This constructor initializes a {@code Builder} object with a null {@code CellStyle}, which means no default
         * stylistic attributes will be applied to the cells created using this builder unless explicitly set later.
         */
        public Builder() {
            this.cellStyle = null;
        }

        /**
         * Sets the value of the cell.
         * @param value the value to be set in the cell.
         * @return the current Builder instance for method chaining.
         */
        public Builder value(Object value) {
            this.value = value;
            return this;
        }

        /**
         * Sets the cell operation to be performed on the cell.
         * @param cellOperation the {@link CellOperation} to be performed on the cell.
         * @return the current Builder instance for method chaining.
         */
        public Builder cellOperation(CellOperation cellOperation) {
            this.cellOperation = cellOperation;
            return this;
        }

        /**
         * Builds a {@link CellAttribute} instance with the specified attributes.
         * @return a new {@link CellAttribute} instance with the attributes set in the builder.
         */
        public CellAttribute build() {
            return new CellAttribute(this);
        }

    }
}
