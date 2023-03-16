package com.dasilvadeveloper.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;

public class XlsxFactory {

    private XSSFWorkbook workbook = new XSSFWorkbook();
    private Sheet sheet;
    private Layout layout;
    private ArrayList<?> dataModel;
    private Row headerRow;
    private String[] headerColumnAttributes;
    private String[] headerTitles;
    private String dataModelClassName;
    private CellStyle headerStyle;
    private XSSFFont font;
    private CellStyle style;

    //constructors

    public XlsxFactory() {
    }

    public XlsxFactory(Layout layout) {
        this.layout = layout;

    }

    /**
     * builds a xls file
     *
     * @param to output path
     */
    public void build(String to, String fileName, boolean includeDateTime) throws IOException, NoSuchFieldException, IllegalAccessException {
        // transform file name to lower case and replaces spaces to _
        fileName = fileName.toLowerCase().replace(' ', '_');

        this.sheet = this.workbook.createSheet(fileName);

        // create header
        createHeader(this.workbook, this.sheet, this.headerColumnAttributes);

        // fill data
        createBody(dataModel);

        // get date time
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("uuuuMMddHHmmss");
        String dateStr = dtf.format(now);

        // output
        // add date to file name?
        if(includeDateTime) fileName += "_" + dateStr + ".xlsx";
        else fileName += ".xlsx";

        File currDir = new File("");
        String path = currDir.getAbsolutePath() + to;
        String fileLocation = path + fileName;

        // create directory if not exists
        new File(path).mkdirs();

        // create the file and close the stream
        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

    }

    private void createBody(ArrayList<?> dataModel) throws NoSuchFieldException, IllegalAccessException {
        // create simple cell style
        this.style = workbook.createCellStyle();

        style.setWrapText(false);

        // iterate between each row
        for (int i = 0; i < dataModel.size(); i++) {

            // create rows with data
            Row row = sheet.createRow(i + headerRow.getRowNum() +1);
            row.setHeight((short) (headerRow.getHeight() * 1.02));

            // fill cells of each row with data
            for (int j = 0; j < headerColumnAttributes.length; j++) {
                Cell cell = row.createCell(j);

                Object o = dataModel.get(i);

                Class<?> c = o.getClass();

                Field f = c.getDeclaredField(headerColumnAttributes[j]);
                f.setAccessible(true);


                // fill with different data depending on the index value

                cell.setCellValue(f.get(o).toString());

                cell.setCellStyle(style);
            }


        }
    }

    private void createHeader(XSSFWorkbook workbook, Sheet sheet, String[] headerColumnNames) {

        for (int i = 0; i < headerTitles.length; i++) {
            // set columns width
            sheet.setColumnWidth(i, 10000);
        }

        // create the header row
        this.headerRow = sheet.createRow(0);
        headerRow.setHeight((short) (headerRow.getHeight() * 2));

        // create header cell style
        this.headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // set header font style
        this.font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        // fill the header with data
        for (int i = 0; i < headerTitles.length; i++) {

            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headerTitles[i]);
            headerCell.setCellStyle(headerStyle);

        }

    }


    // util classes
    public enum Layout {
        SIMPLE
    }

    // getter & setters
    public Layout getLayout() {
        return layout;
    }

    public void setLayout(Layout layout) {
        this.layout = layout;
    }

    public Object getDataModel() {
        return dataModel.get(0);
    }

    public void setDataModel(ArrayList<?> dataModel) {
        this.dataModel = dataModel;
        this.dataModelClassName = dataModel.get(0).getClass().getSimpleName();
    }

    public String[] getHeaderColumnAttributes() {
        return headerColumnAttributes;
    }

    /**
     * each attribute must exactly match the attribute of the object
     * @param headerColumnAttributes array of attr names
     */
    public void setHeaderColumnAttributes(String[] headerColumnAttributes) {
        this.headerColumnAttributes = headerColumnAttributes;
    }

    public String[] getHeaderTitles() {
        return headerTitles;
    }

    public void setHeaderTitles(String[] headerTitles) {
        this.headerTitles = headerTitles;
    }
}
