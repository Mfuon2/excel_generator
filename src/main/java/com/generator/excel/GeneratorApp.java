package com.generator.excel;

import com.generator.excel.model.LogModel;
import com.generator.excel.service.ExcelGeneratorService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class GeneratorApp {

    public static  void main(String[] args)
    {
        generateExcel();
    }

    public static void generateExcel() {

        List<LogModel> list= new ArrayList<>();
        LogModel log = new LogModel(1L,"John", BigDecimal.valueOf(35000),BigDecimal.valueOf(4.3));
        LogModel log2 = new LogModel(3L,"Leonard", BigDecimal.valueOf(35000),BigDecimal.valueOf(4.3));
        LogModel log3 = new LogModel(2L,"Mfuon", BigDecimal.valueOf(40000),BigDecimal.valueOf(3.4));

        list.add(log);
        list.add(log2);
        list.add(log3);

        try
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFSheet sheet = workbook.createSheet("sheet1");// creating a blank sheet
            int rownum = 1;
            Row row1 = sheet.createRow(0);

            createHeader(row1,workbook);
            for (LogModel me : list)
            {
                Row row = sheet.createRow(rownum++);
                createList(me, row, workbook);

            }

            FileOutputStream out = new FileOutputStream("./target/"+ new Date() +" PREMIUM RATE CHANGE.xlsx"); // file name with path
            workbook.write(out);
            out.close();

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }

    public static void createList(LogModel model, Row row, XSSFWorkbook workbook) {

        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyle numberCellStyle = workbook.createCellStyle();
        numberCellStyle.setAlignment((short) 3);
        numberCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));

        Cell cell = row.createCell(0);
        cell.setCellValue(model.getLogId());

        cell = row.createCell(1);
        cell.setCellValue(model.getName());

        cell = row.createCell(2);
        cell.setCellStyle(numberCellStyle);
        cell.setCellValue(model.getPremLimit().toString());

        cell = row.createCell(3);
        cell.setCellStyle(numberCellStyle);
        cell.setCellValue( model.getRate().toPlainString());

    }

    public static void createHeader(Row row, XSSFWorkbook workbook) {
        String[] columns = {"ID","NAME","LIMIT","RATE"};
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.BLUE.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setAlignment((short) 3);
        // Header
        for (int col = 0; col < columns.length; col++) {
            Cell cell = row.createCell(col);
            cell.setCellValue(columns[col]);
            cell.setCellStyle(headerCellStyle);
        }

    }
}
