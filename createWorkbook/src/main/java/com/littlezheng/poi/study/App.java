package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

    public static void main(String[] args) {
        File dir = getOutDir();
        try (Workbook wb = new HSSFWorkbook()) {
            wb.write(new FileOutputStream(new File(dir, "1.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }

        try (Workbook wb2 = new XSSFWorkbook()) {
            wb2.write(new FileOutputStream(new File(dir, "2.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public static File getOutDir() {
        File path = new File("out");
        if (!path.exists()) {
            path.mkdirs();
        }
        return path;
    }
    
}
