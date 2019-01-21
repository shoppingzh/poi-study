package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void lockSheet() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet("my sheet");
            Cell c = s.createRow(3).createCell(3);
            c.setCellValue("Hello");
            CellStyle cs = wb.createCellStyle();
            cs.setLocked(false);
            c.setCellStyle(cs);
            s.protectSheet("123");
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
