package io.github.vatisteve.utils.excel;

import io.github.vatisteve.utils.excel.common.ElementIdentifier;
import io.github.vatisteve.utils.excel.common.ExcelElement;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("ElementNotFoundException")
class ElementNotFoundExceptionTest {

    @Test
    @DisplayName("exposes element, identifier and position")
    void exposesDetails() {
        ElementNotFoundException ex =
                new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.NAME, "foo");
        assertEquals(ExcelElement.SHEET, ex.getElement());
        assertEquals(ElementIdentifier.NAME, ex.getIdentifier());
        assertArrayEquals(new Object[]{"foo"}, ex.getPosition());
    }

    @Test
    @DisplayName("builds a descriptive message")
    void buildsMessage() {
        ElementNotFoundException single =
                new ElementNotFoundException(ExcelElement.SHEET, ElementIdentifier.NAME, "foo");
        assertEquals("There is no SHEET-NAME with '[foo]'", single.getMessage());

        ElementNotFoundException cell =
                new ElementNotFoundException(ExcelElement.CELL, ElementIdentifier.POSITION, 2, 3);
        assertEquals("There is no CELL-POSITION with '[2, 3]'", cell.getMessage());
    }
}
