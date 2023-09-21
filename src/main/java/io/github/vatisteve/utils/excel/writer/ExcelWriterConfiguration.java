package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.CellStyle;

import java.time.ZoneId;

public interface ExcelWriterConfiguration {
    String getDateTimeFormat();
    ZoneId getZoneId();
    CellStyle getCellStyle();
    boolean isWithExcelHeader();
    Object[] getHeaderData();
    String getNumberFormat();
    short getRowHeight();

    class DefaultConfiguration implements ExcelWriterConfiguration {

        @Override
        public String getDateTimeFormat() {
            return "yyyy/MM/dd HH:mm";
        }

        @Override
        public ZoneId getZoneId() {
            return ZoneId.systemDefault();
        }

        @Override
        public CellStyle getCellStyle() {
            return null;
        }

        @Override
        public boolean isWithExcelHeader() {
            return false;
        }

        @Override
        public Object[] getHeaderData() {
            return new Object[] {};
        }

        @Override
        public String getNumberFormat() {
            return null;
        }

        @Override
        public short getRowHeight() {
            return -1; // auto size
        }
    }
}
