package com.gaolei.excelpoi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 16:59
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: 读取Excel数据
 */
public class ReadExcel {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("Test.xlsx"));
            //使用Test.xlsx文件创建工作簿对象
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //获取第一个sheet内容
            XSSFSheet sheet = workbook.getSheetAt(0);
            // 逐行遍历
            Iterator<Row> rowIterable = sheet.iterator();
            while (rowIterable.hasNext()){

                Row row = rowIterable.next();
                // 逐列遍历
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
//                        case Cell.CELL_TYPE_NUMERIC:
//                            System.out.println(cell.getNumericCellValue() + "t");
//                            break;
//                        case Cell.CELL_TYPE_STRING:
//                            System.out.println(cell.getStringCellValue() + "t");
//                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
