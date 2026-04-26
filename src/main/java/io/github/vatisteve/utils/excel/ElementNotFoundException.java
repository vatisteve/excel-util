package io.github.vatisteve.utils.excel;

import io.github.vatisteve.utils.excel.common.ElementIdentifier;
import io.github.vatisteve.utils.excel.common.ExcelElement;

import java.util.Arrays;

/**
 * Exception thrown when an expected Excel element cannot be found during operations.
 *<p>
 * This class is intended to provide detailed information about the missing element,
 * including its type, identifier, and position within the Excel structure.
 */
public final class ElementNotFoundException extends Exception {

    private static final long serialVersionUID = 1017632711482916265L;

    /**
     * The type of Excel element that was not found.
     */
    private final ExcelElement element;
    /**
     * The method of identification used to locate the element.
     */
    private final ElementIdentifier identifier;
    private final transient Object[] position;

    /**
     * Constructs an instance of {@code ElementNotFoundException} to indicate that a specific Excel
     * element could not be located during processing.
     *
     * @param element the type of Excel element that was not found (e.g., SHEET, ROW, COLUMN, CELL)
     * @param identifier the method of identification used to locate the element (e.g., NAME, POSITION)
     * @param position the additional parameters or location details used during the
     *                 element lookup
     */
    public ElementNotFoundException(ExcelElement element, ElementIdentifier identifier, Object... position) {
        super(String.format("There is no %s-%s with '%s'", element.name(), identifier, Arrays.asList(position)));
        this.element = element;
        this.identifier = identifier;
        this.position = position;
    }

    /**
     * Get the type of Excel element that was not found.
     * @return the type of Excel element that was not found.
     */
    public ExcelElement getElement() {
        return element;
    }

    /**
     * Get the method of identification used to locate the element.
     * @return the method of identification used to locate the element.
     */
    public ElementIdentifier getIdentifier() {
        return identifier;
    }

    /**
     * Get the additional parameters or location details that were used during the element lookup.
     * @return the additional parameters or location details used during the element lookup.
     */
    public Object[] getPosition() {
        return position;
    }
}
