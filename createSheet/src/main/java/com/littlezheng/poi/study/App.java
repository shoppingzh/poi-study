package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Test;

public class App {

    @Test
    public void createDefaultSheet() {
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet s1 = wb.createSheet();
            System.out.println("sheet name: " + s1.getSheetName());
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void createNamedSheet() {
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet s1 = wb.createSheet("mySheet");
            System.out.println("sheet name: " + s1.getSheetName());
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void createManySheets() {
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet s1 = wb.createSheet("mySheet1");
            Sheet s2 = wb.createSheet();
            Sheet s3 = wb.createSheet("mySheet3");
            System.out.println("sheet1 name: " + s1.getSheetName());
            System.out.println("sheet2 name: " + s2.getSheetName());
            System.out.println("sheet3 name: " + s3.getSheetName());
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 创建非法字符sheet
     */
    @Test
    public void createIllegalCharacterSheet() {
        try (Workbook wb = new HSSFWorkbook()) {
//            wb.createSheet("sheet*1");    // no
//            wb.createSheet("sheet-1");    // yes
//            wb.createSheet("sheet:1");    // no
//            wb.createSheet("sheet/1");    // no
//            wb.createSheet("sheet(1)");   // yes
//            wb.createSheet("sheet[1]");   // no
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 创建名称安全的sheet
     */
    @Test
    public void createSafeNameSheet() {
        try (Workbook wb = new HSSFWorkbook()) {
            wb.createSheet(WorkbookUtil.createSafeSheetName("sheet[1]")); // 默认以空格替换掉非法字符
            wb.createSheet(WorkbookUtil.createSafeSheetName("sheet*2", '-')); // 指定替换字符
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 将当前workbook写出到out文件夹，文件名为调用该方法的方法名
     * 
     * @param wb                workbook
     * @throws IOException      写出异常
     */
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
