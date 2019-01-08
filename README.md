# poi-study
![logo](images/1.jpg)



## 1. 创建WorkBook
```java
try (Workbook wb = new HSSFWorkbook()) {
	wb.write(new FileOutputStream("d:/1.xls"));
} catch (IOException e) {
	e.printStackTrace();
}

try (Workbook wb2 = new XSSFWorkbook()) {
	wb2.write(new FileOutputStream("d:/2.xlsx"));
} catch (IOException e) {
	e.printStackTrace();
}
```

**注**：创建的excel文件是无法打开的，因为一个sheet都没有。效果如下图所示：

##### 1.xls
![1.xls](images/2.jpg)

##### 2.xlsx
![2.xlsx](images/3.jpg)



## 2. 创建Sheet
```java
try (Workbook wb = new HSSFWorkbook()) {
    Sheet s1 = wb.createSheet();
    System.out.println("sheet name: " + s1.getSheetName());
    write(wb);
} catch (IOException e) {
    e.printStackTrace();
}
```
上述是最简单的创建sheet的例程，控制台打印如下：
> sheet name: Sheet0

默认以Sheet[Num]为名称，Num为当前的sheet的序号，从0开始。除此之外，还可以手动指定sheet名称，如：
```java
try (Workbook wb = new HSSFWorkbook()) {
    Sheet s1 = wb.createSheet("mySheet");
    System.out.println("sheet name: " + s1.getSheetName());
    write(wb);
} catch (IOException e) {
    e.printStackTrace();
}
```
>sheet name: mySheet

**注意**：并非所有字符都可以作为sheet name，如***sheet[1]***这样的名称就不允许作为sheet的名称，如果对非法字符不清楚或无法规避，则可使用下述的生成sheet name方法：

```java
try (Workbook wb = new HSSFWorkbook()) {
    wb.createSheet(WorkbookUtil.createSafeSheetName("sheet[1]"));  
    wb.createSheet(WorkbookUtil.createSafeSheetName("sheet*2", '-'));
    write(wb);
} catch (IOException e) {
    e.printStackTrace();
}
```
默认将以空格替换掉非法字符，如果需要自定义替换字符，则调用**WorkbookUtil**的重载方法*createSafeSheetName(sheetName, replaceChar)*



## 3. 创建Cell(单元格)
```java
try(Workbook wb = WorkbookFactory.create(false)){
    Sheet s = wb.createSheet("mySheet");
    Row r1 = s.createRow(1);
    Cell r1c0 = r1.createCell(0);
    r1c0.setCellValue("Hello, world!");
    
    Cell r1c1 = r1.createCell(1);
    r1c1.setCellValue(true);
    wb.write(new FileOutputStream("out/1.xls"));
}catch (IOException e) {
    e.printStackTrace();
}
```
单元格(Cell)对象由行(Row)对象持有，在创建单元格之前，先使用Sheet对象创建一个行，然后使用Row创建一个单元格，通过**setCellValue**方法设置单元格的值，可设置的类型有：
* **double**：数值类型，如1.25
* **boolean**：布尔类型，如true
* **String**：字符串类型，如"hi"
* **Calendar/Date**：日期类型
* **RichTextString**：富文本类型

直接使用**setCellValue(val)**方法设置单元格的值，各类型的表现如下图所示(office 2010)：
![cell vlaue perform](images/5.jpg)

通过上图可以发现各类型的值在表现上有以下两个问题：
1. 横向对齐方式：数值类型的默认右对齐，布尔类型居中，字符串左对齐（**如果不指定类型，则默认使用org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL的对齐风格**）
2. 日期值显示错误：显示的日期值实际为数值，这是因为单元格没有进行数据格式化，接下来将演示如何创建一个日期类型的单元格
### 创建日期Cell
```java
try (Workbook wb = WorkbookFactory.create(false)) {
    Sheet s = wb.createSheet("mySheet");
    Row r1 = s.createRow(1);
    Cell r1c0 = r1.createCell(0);
    r1c0.setCellValue(new Date());

    CellStyle cs = wb.createCellStyle();
    cs.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy h:mm"));
    r1c0.setCellStyle(cs);
    write(wb);
} catch (IOException e) {
    e.printStackTrace();
}
```
对于日期类型这种特殊格式的字符串，需要设置单元格的样式，并将格式化的方法告知CellStyle对象，就可以将普通字符串格式化为特殊字符串了。
**关键代码**：

```java
CellStyle cs = wb.createCellStyle();
cs.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy h:mm"));
```

**BuiltinFormats.java中定义的格式化字符串：**

```java
private final static String[] _formats = {
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "\"$\"#,##0_);(\"$\"#,##0)",
        "\"$\"#,##0_);[Red](\"$\"#,##0)",
        "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        "0%",
        "0.00%",
        "0.00E+00",
        "# ?/?",
        "# ??/??",
        "m/d/yy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "h:mm AM/PM",
        "h:mm:ss AM/PM",
        "h:mm",
        "h:mm:ss",
        "m/d/yy h:mm",

        // 0x17 - 0x24 reserved for international and undocumented
        // TODO - one junit relies on these values which seems incorrect
        "reserved-0x17",
        "reserved-0x18",
        "reserved-0x19",
        "reserved-0x1A",
        "reserved-0x1B",
        "reserved-0x1C",
        "reserved-0x1D",
        "reserved-0x1E",
        "reserved-0x1F",
        "reserved-0x20",
        "reserved-0x21",
        "reserved-0x22",
        "reserved-0x23",
        "reserved-0x24",
        
        "#,##0_);(#,##0)",
        "#,##0_);[Red](#,##0)",
        "#,##0.00_);(#,##0.00)",
        "#,##0.00_);[Red](#,##0.00)",
		"_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        "mm:ss",
        "[h]:mm:ss",
        "mm:ss.0",
        "##0.0E+0",
        "@"
	};
```



## 4. 设置单元格格式




## 5. 单元格对齐
```java
try (Workbook wb = WorkbookFactory.create(false)) {
    Sheet s = wb.createSheet("mySheet");
    s.setDefaultColumnWidth(20);
    Row r0 = s.createRow(0);
    r0.setHeightInPoints(30f);
    Cell r0c0 = r0.createCell(0);
    r0c0.setCellValue("center text");
    CellStyle cs = wb.createCellStyle();
    cs.setAlignment(HorizontalAlignment.CENTER);
    cs.setVerticalAlignment(VerticalAlignment.CENTER);
    r0c0.setCellStyle(cs);
    write(wb);
} catch (IOException e) {
    e.printStackTrace();
}
```
以上代码设置单元格的水平、垂直对齐方式为居中对齐，该操作通过**CellStyle.setAlignment(HorizontalAlignment)**与**CellStyle.setVerticalAlignment(VerticalAlignment)**两个方法完成。

excel中，水平对齐有以下对齐方式：
1. 常规(General)：自动对齐，数字类型居右，字符串居左，布尔类型居中
2. 居左(Left)：靠左对齐，**可设置缩进**
3. 居中(Center)：居中对齐
4. 居右(Left)：靠右对齐，**可设置缩进**
5. 填充(Fill)：将内容在单元格的水平方向填充
6. 两端对齐(Justify)：除了最后一行，其它每行文本左右对齐（**注：没有试出效果**）
7. 跨列居中(Center selection)：与居中不同，跨列居中的单元格在单元格的宽度不够时，会跨越到其它列显示
8. 分散对齐(Distrubuted)：将单元格中的每个**词**均匀地分布在单元格中，单元格宽度改变，各个词之间的间隔动态变化

以上八中对齐方式在excel中的表现如下图：
![](images/6.jpg)



垂直对齐方式：
1. 居上(Top)：内容贴近单元格顶部
2. 居中(Center)：内容在单元格垂直方向居中
3. 居下(Bottom)：内容贴近单元格底部
4. 两端对齐(Justify)：内容在垂直方向均匀分布
5. 分散对齐(Distrubuted)：与两端对齐效果相同

以上5中对齐方式在excel中的表现如下图：
![](images/7.jpg)



## 6. 为单元格设置边框
```java
try(Workbook wb = WorkbookFactory.create(true)){
    Sheet s = wb.createSheet("my sheet");
    Row r = s.createRow(2);
    Cell c = r.createCell(5);
    r.setHeightInPoints(50f);
    c.setCellValue("hello");
    
    CellStyle cs = wb.createCellStyle();
    c.setCellStyle(cs);
    cs.setBorderTop(BorderStyle.THIN);
    cs.setTopBorderColor(IndexedColors.RED.getIndex());
    cs.setBorderRight(BorderStyle.MEDIUM_DASHED);
    cs.setRightBorderColor(IndexedColors.BLUE.getIndex());
    cs.setBorderBottom(BorderStyle.DOTTED);
    cs.setBottomBorderColor(IndexedColors.PINK.getIndex());
    cs.setBorderLeft(BorderStyle.DOUBLE);
    cs.setLeftBorderColor(IndexedColors.YELLOW.getIndex());
    
    write(wb);
}catch (IOException e) {
    e.printStackTrace();
}
```
边框的设置有三要素：位置(position)，粗细(width)，颜色(color)，poi中设置边框样式的示意图如下：
![](images/8.jpg)

在excel中设置边框的方式如下所示：
![](images/9.jpg)
可见，在excel中设置边框的选择更多样，可以设置斜的边框（严格意义上说这应该不算边框）。



## 7. 设置单元格背景
```java

```
