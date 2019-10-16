package com.gaolei.excelpoi.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Random;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 18:34
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: 格式化Excel样式
 */
public class FomatExcel {
    public static final Integer NUM = 100;
    public static void main(String[] args) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
//        XSSFSheet sheet = workbook.createSheet("单元格样式");
//        // formatByValue(sheet);
//        FileOutputStream out = new FileOutputStream("styleDemo.xlsx");
//        XSSFSheet sheet = workbook.createSheet("交替行着色");
//        formatByColor(sheet);
//        FileOutputStream out = new FileOutputStream("colorDemo.xlsx");

        XSSFSheet sheet = workbook.createSheet("到期时间");
        expiryInNext30Days(sheet);
        FileOutputStream out = new FileOutputStream("Date.xlsx");
        workbook.write(out);
        out.close();
    }

    public static  void  formatByValue(Sheet sheet){
        Random random = new Random();
        for (int i = 0; i < NUM; i++) {
            sheet.createRow(i).createCell(0).setCellValue(random.nextInt(100));
        }
        // 获取格式化对象
        SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
        //设置格式化条件，条件1
        ConditionalFormattingRule rule1 = conditionalFormatting.createConditionalFormattingRule(ComparisonOperator.GT,"70");
        PatternFormatting patternFormatting1  = rule1.createPatternFormatting();
        patternFormatting1 .setFillBackgroundColor(IndexedColors.BLUE.index);
        patternFormatting1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // 条件2
        ConditionalFormattingRule rule2 = conditionalFormatting.createConditionalFormattingRule(ComparisonOperator.LT,"50");
        PatternFormatting patternFormatting2  = rule2.createPatternFormatting();
        patternFormatting2 .setFillBackgroundColor(IndexedColors.RED.index);
        patternFormatting2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // 列格式范围
        CellRangeAddress[] range = {
                CellRangeAddress.valueOf("A1:A100")
        };
        conditionalFormatting.addConditionalFormatting(range,rule1,rule2);
    }

    public static void formatByColor(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule);
        sheet.createRow(0).createCell(1).setCellValue("交替行变色，绿色填充");
        sheet.createRow(1).createCell(1).setCellValue("条件：MOD(ROW(),2)");
    }

    public static void expiryInNext30Days(Sheet sheet)
    {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setDataFormat((short)BuiltinFormats.getBuiltinFormat("d-mmm"));

        sheet.createRow(0).createCell(0).setCellValue("日期");
        sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
        sheet.createRow(2).createCell(0).setCellFormula("A2+1");
        sheet.createRow(3).createCell(0).setCellFormula("A3+1");

        for(int rownum = 1; rownum <= 3; rownum++) sheet.getRow(rownum).getCell(0).setCellStyle(style);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A4")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(0).createCell(1).setCellValue("样式将在30后到期");
    }
}
