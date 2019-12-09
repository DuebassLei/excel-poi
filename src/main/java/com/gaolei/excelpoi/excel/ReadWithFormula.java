package com.gaolei.excelpoi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

import java.util.Iterator;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 17:57
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: 读取带有计算公式formula的Excel
 */
public class ReadWithFormula {

    public static void main(String[] args) {
//        XSSFWorkbook workbook = new XSSFWorkbook();
//
//        try {
//            FileInputStream file  = new FileInputStream(new File("formulaTest.xlsx"));
//            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//
//            XSSFSheet sheet = workbook.getSheetAt(0);
//            // 遍历Row
//            Iterator<Row> rowIterator = sheet.iterator();
//            while (rowIterator.hasNext()){
//                Row row = rowIterator.next();
//                // 遍历Cell
//                Iterator<Cell> cellIterator = row.cellIterator();
//                while (cellIterator.hasNext()){
//                    Cell cell = cellIterator.next();
//                    switch (evaluator.evaluateInCell(cell).getCellType()){
//                        case Cell.CELL_TYPE_NUMERIC:
//                            System.out.println(cell.getNumericCellValue() + "tt");
//                            break;
//                        case Cell.CELL_TYPE_STRING:
//                            System.out.println(cell.getStringCellValue() + "tt");
//                            break;
//                        case Cell.CELL_TYPE_FORMULA:
//                            break;
//                    }
//                }
//                System.out.println("");
//            }
//            file.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
    }
}
