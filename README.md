# Apache POI对Excel的基本操作
## code目录
- com.gaolei.excelpoi
    - excel
        - a. WriterExcel 插入数据
        - b. ReadExcel 读取数据
        - c. FormulaExcel 插入数据（公式计算）
        - d. ReadWithFormula 读取数据（公式计算）
## 环境
- IntelliJ IDEA 2018.2
- JDK 1.8
- SpringBoot 2.1.9.RELEASE
- POI 3.9
## Maven 依赖
```xml
    <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.9</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.9</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml-schemas</artifactId>
            <version>3.9</version>
        </dependency>
    </dependencies>
```
### 插入数据Excel
```java
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
```

## 读取数据Excel
```java
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
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println(cell.getStringCellValue() + "t");
                            break;
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
```

## 插入带公式计算Excel
```java
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
```

## 读取带公式计算的Excel
```java
public class ReadWithFormula {

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        try {
            FileInputStream file  = new FileInputStream(new File("formulaTest.xlsx"));
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            XSSFSheet sheet = workbook.getSheetAt(0);
            // 遍历Row
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                // 遍历Cell
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (evaluator.evaluateInCell(cell).getCellType()){
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "tt");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println(cell.getStringCellValue() + "tt");
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            break;
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
```

## 格式化表格
### 按值大小
```java
public class FomatExcel {
    public static final Integer NUM = 100;
    public static void main(String[] args) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("单元格样式");
        formatByValue(sheet);
        FileOutputStream out = new FileOutputStream("styleDemo.xlsx");
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
}
```

### 交替行变色
```java
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
}
```
### 设置到期时间
```java
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
```

## 查看更多
[DuebassLei的CSDN小窝](https://blog.csdn.net/m0_37903882)

