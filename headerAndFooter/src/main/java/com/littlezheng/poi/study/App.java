package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void header() {
        try (Workbook wb = WorkbookFactory.create(true)) {

            Sheet s = wb.createSheet("my sheet");
            s.createRow(3).createCell(4).setCellValue("Hello");
            Header h = s.getHeader();
            h.setCenter("Header center");

            write(wb);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void footer() {
        try (Workbook wb = WorkbookFactory.create(true)) {

            Sheet s = wb.createSheet("my sheet");
            s.createRow(3).createCell(4).setCellValue("Hello");
            Footer f = s.getFooter();
            f.setRight("Fotter right");
            
            
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
