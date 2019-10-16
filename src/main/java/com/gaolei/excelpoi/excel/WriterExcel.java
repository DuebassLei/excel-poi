package com.gaolei.excelpoi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 16:27
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: Excel写入数据
 */
public class WriterExcel {
    public static void main(String[] args) {
        // 1. 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 2. 创建工作表
        XSSFSheet sheet = workbook.createSheet("WriterDataTest");
        // 3. 模拟待写入数据
        Map<String,Object[]> data = new TreeMap<>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {1, "Amit", "Shukla"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data.put("4", new Object[] {3, "John", "Adwards"});
        data.put("5", new Object[] {4, "Brian", "Schultz"});
        //4. 遍历数据写入表中
        Set<String> keySet = data.keySet();
        int rowNum = 0;
        for (String key : keySet){
            Row row = sheet.createRow(rowNum++);
            Object [] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj: objArr){
                Cell cell  = row.createCell(cellNum++);
                if (obj instanceof String){
                    cell.setCellValue((String)obj);
                }else if(obj instanceof Integer){
                    cell.setCellValue((Integer)obj);
                }
            }
        }
        try {
            File file = new File("Test.xlsx");
            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
