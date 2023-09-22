package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

import java.time.ZoneId;

public interface ExcelWriterConfiguration {
    String defaultLocalTimeFormat();
    ZoneId defaultZoneId();
    CellStyle defaultCellStyle();
    boolean withExcelHeader();
    Object[] excelHeaderData();
    short defaultRowHeight();

    class DefaultConfiguration implements ExcelWriterConfiguration {

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
        public boolean withExcelHeader() {
            return false;
        }

        @Override
        public Object[] excelHeaderData() {
            return new Object[] {};
        }

        @Override
        public short defaultRowHeight() {
            return -1; // auto size
        }
    }
}
