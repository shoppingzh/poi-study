package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {
    
    public static void main(String[] args) {
        System.out.println(days("2019-01-03", "1900-12-31"));
    }

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

    @Test
    public void unFormatCell() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet("my sheet");
            Row r0 = s.createRow(0);
            Calendar c = Calendar.getInstance();
            c.set(2019, 0, 3, 16, 20, 15);
            c.set(Calendar.MILLISECOND, 0);
            r0.createCell(0).setCellValue(c.getTime());

            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void write(Workbook wb) throws IOException {
        String fileName = new Exception().getStackTrace()[1].getMethodName()
                + (wb instanceof HSSFWorkbook ? ".xls" : ".xlsx");
        wb.write(new FileOutputStream(new File(getOutDir(), fileName)));
    }

    public static File getOutDir() {
        File path = new File("out");
        if (!path.exists()) {
            path.mkdirs();
        }
        return path;
    }

}
