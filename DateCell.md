## 奇怪的数字

```java
try(Workbook wb = WorkbookFactory.create(true)){
    Sheet s = wb.createSheet("my sheet");
    Row r0 = s.createRow(0);
    Calendar c = Calendar.getInstance();
    c.set(2019, 0, 3, 16, 20, 15);
    c.set(Calendar.MILLISECOND, 0);
    r0.createCell(0).setCellValue(c.getTime());
    
    write(wb);
}catch(IOException e){
    e.printStackTrace();
}
```
以上代码为第一行的第一个单元格设置了2019-01-03 16:20:15.000(Calendar的1月为MONTH=0)的日期值，生成的效果如下图所示：
![date cell](images/dateCell/1.jpg)

为什么生成的excel中单元格的值非但不是一个日期值，反而是一个很奇怪的数值？查看**org.apache.poi.xssf.usermodel.XSSFCell.setCellValue(Date)**源码：
```java
@Override
public void setCellValue(Date value) {
    if(value == null) {
        setCellType(CellType.BLANK);
        return;
    }

    boolean date1904 = getSheet().getWorkbook().isDate1904();
    setCellValue(DateUtil.getExcelDate(value, date1904));
}
```
`DateUtil.getExcelDate(value, date1904)`返回的是一个double值，进入该方法：
```java
public static double getExcelDate(Date date, boolean use1904windowing) {
    Calendar calStart = LocaleUtil.getLocaleCalendar();
    calStart.setTime(date);   // If date includes hours, minutes, and seconds, set them to 0
    return internalGetExcelDate(calStart, use1904windowing);
}
```
前两行代码的作用是获取传入的时间值的Calendar对象，关键代码在最后一行的**internalGetExcelDate(Caldendar, boolean)**方法，进入该方法：
```java
private static double internalGetExcelDate(Calendar date, boolean use1904windowing) {
    // check
    
    // Because of daylight time saving we cannot use
    //     date.getTime() - calStart.getTimeInMillis()
    // as the difference in milliseconds between 00:00 and 04:00
    // can be 3, 4 or 5 hours but Excel expects it to always
    // be 4 hours.
    // E.g. 2004-03-28 04:00 CEST - 2004-03-28 00:00 CET is 3 hours
    // and 2004-10-31 04:00 CET - 2004-10-31 00:00 CEST is 5 hours
    double fraction = (((date.get(Calendar.HOUR_OF_DAY) * 60
                         + date.get(Calendar.MINUTE)
                        ) * 60 + date.get(Calendar.SECOND)
                       ) * 1000 + date.get(Calendar.MILLISECOND)
                      ) / ( double ) DAY_MILLISECONDS;
    Calendar calStart = dayStart(date);

    double value = fraction + absoluteDay(calStart, use1904windowing);

    if (!use1904windowing && value >= 60) {
        value++;
    } else if (use1904windowing) {
        value--;
    }

    return value;
}
```

**fraction**变量由当前时间的毫秒数除以一天中总的毫秒数得到，正如变量名的中文含义**分数**，该变量实际标识了当前天在一天中占的比例，如上图生成的值的小数约为0.68073，即说明这一天走完了68%，根据16点相对于一天24个小时的比例可以推测该结果是正确的，现在，让我们拿出计算器来进行以下运算（一天中共有86400000毫秒）：

> (16 * 60 * 60 * 1000  + 20 * 60 * 1000 + 15 * 1000 ) / 86400000

结果为**0.68072916666**，该结果四舍五入后与生成的单元格中的值**0.68073**刚好吻合！



接下来是整数部分数值了，整数值为**43468**，该值是如何得到的呢？

```java
double value = fraction + absoluteDay(calStart, use1904windowing);
```

这句代码中的**absoluteDay(calStart, use1904windowing)**就是求解整数部分的值，通过方法名可知该方法求的是传入日期与某个日期值之前的差值，该方法如下：

```java
/**
 * Given a Calendar, return the number of days since 1900/12/31.
 *
 * @return days number of days since 1900/12/31
 * @param  cal the Calendar
 * @exception IllegalArgumentException if date is invalid
 */
protected static int absoluteDay(Calendar cal, boolean use1904windowing){
    return cal.get(Calendar.DAY_OF_YEAR)
           + daysInPriorYears(cal.get(Calendar.YEAR), use1904windowing);
}
```
请注意该方法的文档：该方法的返回值是从1900年12月31日到当前时间的天数
