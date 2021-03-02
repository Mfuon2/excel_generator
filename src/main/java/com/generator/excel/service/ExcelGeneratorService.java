package com.generator.excel.service;

import com.generator.excel.model.LogModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ExcelGeneratorService implements IExcelGeneratorService{
    @Override
    public void generateExcel() {

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
            int rownum = 0;
            for (LogModel me : list)
            {
                Row row = sheet.createRow(rownum++);
                createList(me, row);

            }

            FileOutputStream out = new FileOutputStream("NewFile "+ 56 +".xlsx"); // file name with path
            workbook.write(out);
            out.close();

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }

    public static void createList(LogModel model, Row row) {

        Cell cell = row.createCell(0);
        cell.setCellValue(model.getLogId());

        cell = row.createCell(1);
        cell.setCellValue(model.getName());

        cell = row.createCell(2);
        cell.setCellValue(model.getPremLimit().toString());

        cell = row.createCell(3);
        cell.setCellValue(model.getRate().toString());

    }
}
