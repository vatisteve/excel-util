package io.github.vatisteve.utils.excel.writer;

import org.apache.poi.ss.usermodel.*;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Date;

public class SampleDomain implements Serializable {

    private static final long serialVersionUID = 3225014967364984271L;

    private String segment;
    private String country;
    private String product;
    private String discountBand;
    private double unitsSold;
    private BigDecimal manufacturing;
    private BigDecimal salePrice;
    private BigDecimal grossSales;
    private BigDecimal discount;
    private BigDecimal sales;
    private BigDecimal cogs;
    private BigDecimal profit;
    private Date date;
    private byte monthNumber;
    private String monthName;
    private int year;

    public SampleDomain() {}

    public void setSegment(String segment) {
        this.segment = segment;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public void setDiscountBand(String discountBand) {
        this.discountBand = discountBand;
    }

    public void setUnitsSold(double unitsSold) {
        this.unitsSold = unitsSold;
    }

    public void setManufacturing(BigDecimal manufacturing) {
        this.manufacturing = manufacturing;
    }

    public void setSalePrice(BigDecimal salePrice) {
        this.salePrice = salePrice;
    }

    public void setGrossSales(BigDecimal grossSales) {
        this.grossSales = grossSales;
    }

    public void setDiscount(BigDecimal discount) {
        this.discount = discount;
    }

    public void setSales(BigDecimal sales) {
        this.sales = sales;
    }

    public void setCogs(BigDecimal cogs) {
        this.cogs = cogs;
    }

    public void setProfit(BigDecimal profit) {
        this.profit = profit;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public void setMonthNumber(byte monthNumber) {
        this.monthNumber = monthNumber;
    }

    public void setMonthName(String monthName) {
        this.monthName = monthName;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public void writeDataTo(ExcelWriter writer) {
        writer.addCell(segment);
        writer.addCell(country);
        writer.addCell(product);
        writer.addCell(discountBand);
        writer.addCell("unitsSold");
        CellStyle moneyStyle = writer.getWorkbook().createCellStyle();
        moneyStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        moneyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        moneyStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"));
        Font font = writer.getWorkbook().createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        moneyStyle.setFont(font);
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(manufacturing).build());
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(salePrice).build());
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(grossSales.doubleValue()).build()); // --> in this data format
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(discount).build());
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(sales).build());
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(cogs).build());
        writer.addCell(new CellAttribute.Builder(moneyStyle).value(profit).build());
        CellStyle dateStyle = writer.getWorkbook().createCellStyle();
        DataFormat format = writer.getWorkbook().createDataFormat();
        dateStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        dateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        dateStyle.setDataFormat(format.getFormat("dd/MM/yyyy"));
        writer.addCell(date, dateStyle);
        writer.addCell(monthNumber);
        writer.addCell(monthName);
        writer.addCell(year);
    }

}
