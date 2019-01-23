package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void freeze() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet();
            for (int i = 1; i <= 200; i++) {
                Row r = s.createRow(i - 1);
                for (int j = 1; j <= 100; j++) {
                    Cell c = r.createCell(j - 1);
                    c.setCellValue(i + "-" + j);
                }
            }

            // s.createFreezePane(5, 3);
            
            s.createFreezePane(5, 3, 10, 5);
            
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void split() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet();
            for (int i = 1; i <= 200; i++) {
                Row r = s.createRow(i - 1);
                for (int j = 1; j <= 100; j++) {
                    Cell c = r.createCell(j - 1);
                    c.setCellValue(i + "-" + j);
                }
            }
            
            s.createSplitPane(10000, 5000, 10, 20, Sheet.PANE_UPPER_RIGHT);
            
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
