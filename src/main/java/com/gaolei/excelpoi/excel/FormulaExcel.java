package com.gaolei.excelpoi.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 17:20
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: Excel公式计算
 */
public class FormulaExcel {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("CalcSimple");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("箱数");
        header.createCell(1).setCellValue("单价");
        header.createCell(2).setCellValue("个数");
        header.createCell(3).setCellValue("总价格");


        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue(10);
        dataRow.createCell(1).setCellValue(2.5);
        dataRow.createCell(2).setCellValue(10);
        dataRow.createCell(3).setCellFormula("A2*B2*C2");
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File("formulaTest.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Excel with formula cells written successfully");
        } catch (Exception e) {
            e.printStackTrace();
        }
       }
}
