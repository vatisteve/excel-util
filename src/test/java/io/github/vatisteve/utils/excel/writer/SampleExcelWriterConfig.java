package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.*;

import java.time.ZoneId;

public class SampleExcelWriterConfig implements ExcelWriterConfiguration {

//    private ExcelHeader excelHeader;

    @Override
    public String defaultLocalTimeFormat() {
        return null;
    }

    @Override
    public ZoneId defaultZoneId() {
        return ZoneId.of("Asia/Saigon");
    }

    @Override
    public CellStyle defaultCellStyle(Workbook activeWb) {
        CellStyle style = activeWb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBottomBorderColor(IndexedColors.SEA_GREEN.getIndex());
        return style;
    }

    @Override
    public ExcelHeader excelHeader(Workbook activeWb) {
        // directly config from existing active workbook
        CellStyle headerStyle = defaultCellStyle(activeWb);
        headerStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font= activeWb.createFont();
        font.setFontName("Arial");
        font.setColor(IndexedColors.RED.getIndex());
        font.setBold(true);
        font.setItalic(false);
        headerStyle.setFont(font);
        return new ExcelHeader.Builder()
                .headers("Header 1", "Header 2", "Header 3", "And more")
                .style(headerStyle)
                .build();
        // or create a new instance with current excelHeader field
    }

    @Override
    public short defaultRowHeight() {
        return 500;
    }
}
