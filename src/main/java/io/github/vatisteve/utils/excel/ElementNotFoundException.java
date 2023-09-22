package io.github.vatisteve.utils.excel;

import io.github.vatisteve.utils.excel.enumeration.ElementIdentifier;
import io.github.vatisteve.utils.excel.enumeration.ExcelElement;

import java.util.Arrays;

public final class ElementNotFoundException extends Exception {

    private static final long serialVersionUID = 1017632711482916265L;

    private final ExcelElement element;
    private final ElementIdentifier identifier;
    private final transient Object[] position;

    public ElementNotFoundException(ExcelElement element, ElementIdentifier identifier, Object... position) {
        super(String.format("There is no %s-%s with '%s'", element.name(), identifier, Arrays.asList(position)));
        this.element = element;
        this.identifier = identifier;
        this.position = position;
    }

    public ExcelElement getElement() {
        return element;
    }

    public ElementIdentifier getIdentifier() {
        return identifier;
    }

    public Object[] getPosition() {
        return position;
    }
}
