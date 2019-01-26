package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    static final Random rand = new Random();

    @Test
    public void adjust() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet();
            Row r = s.createRow(0);
            for (int i = 0; i < 10; i++) {
                int c = rand.nextInt(20);
                // 设置自适应
                r.createCell(i).setCellValue(randStr(c));
                s.autoSizeColumn(i);
            }
            
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    @Test
    public void adjustXs() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet();
            Row r = s.createRow(0);
            for (int i = 0; i < 10; i++) {
                int c = rand.nextInt(20);
                // 设置自适应
                r.createCell(i).setCellValue(randStr(c));
                s.autoSizeColumn(i);
            }
            
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String randStr(int c) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < c; i++) {
            sb.append("aaa").append(i + 1);
        }
        return sb.toString();
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
