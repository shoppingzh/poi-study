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

为什么生成的excel中单元格的值非但不是一个日期值，反而是一个很奇怪的数值？查看`org.apache.poi.xssf.usermodel.XSSFCell.setCellValue(Date)`源码：
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
前两行代码的作用是获取传入的时间值的Calendar对象，关键代码在最后一行的`internalGetExcelDate(Caldendar, boolean)`方法，进入该方法：
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

这句代码中的`absoluteDay(calStart, use1904windowing)`就是求解整数部分的值，通过方法名可知该方法求的是传入日期与某个日期值之前的差值，该方法如下：

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
请注意该方法的文档：该方法的返回值是从1900年12月31日到当前时间的天数，也就是从1900-12-31到2019-01-03的天数，我们可以百度**日期计算器**使用网络上的日期计算器计算这两个日期之间的差值，也可以自己手写计算差值的函数，我们选择了手写该函数：

```java
public static int days(String from, String to) {
    DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
    try {

        Calendar c1 = Calendar.getInstance();
        c1.setTime(df.parse(from));
        Calendar c2 = Calendar.getInstance();
        c2.setTime(df.parse(to));

        boolean lt = c1.getTimeInMillis() > c2.getTimeInMillis();
        Calendar fromCal = lt ? c2 : c1;
        Calendar toCal = lt ? c1 : c2;

        int total = 0;
        int fromYear = fromCal.get(Calendar.YEAR);
        int toYear = toCal.get(Calendar.YEAR);
        if (Math.abs(fromYear - toYear) > 1) {
            Calendar tmp = Calendar.getInstance();
            reset(tmp);
            for (int year = fromYear + 1; year < toYear; year++) {
                tmp.set(Calendar.YEAR, year);
                int dayOfYear = tmp.getActualMaximum(Calendar.DAY_OF_YEAR);
                total += dayOfYear;
            }
        }
        int fromYearDay = fromCal.get(Calendar.DAY_OF_YEAR);
        int toYearDay = toCal.get(Calendar.DAY_OF_YEAR);
        if (fromYear == toYear) { // 同一年
            total += Math.abs(toYearDay - fromYearDay);
        } else {
            total += (fromCal.getActualMaximum(Calendar.DAY_OF_YEAR) - fromYearDay + toYearDay);
        }

        return lt ? total : -1 * total;
    } catch (ParseException e) {
        e.printStackTrace();
    }

    return 0;
}

private static void reset(Calendar tmp) {
    tmp.set(Calendar.MONTH, 0);
    tmp.set(Calendar.DAY_OF_YEAR, 1);
    tmp.set(Calendar.HOUR, 0);
    tmp.set(Calendar.MINUTE, 0);
    tmp.set(Calendar.SECOND, 0);
    tmp.set(Calendar.MILLISECOND, 0);
}
```
调用该方法：`System.out.println(days("2019-01-03", "1900-12-31"))`，得到的日期差值为：
> 43102

为什么