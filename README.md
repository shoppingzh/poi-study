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

直接使用**setCellValue(val)**方法设置单元格的值，各类型的表现如下图所示(office 2010)：
![cell vlaue perform](images/5.jpg)

通过上图可以发现各类型的值在表现上有以下两个问题：
1. 横向对齐方式：数值类型的默认右对齐，布尔类型居中，字符串左对齐
2. 日期值显示错误：显示的日期值实际为数值
