package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void background(){
        try(Workbook wb = WorkbookFactory.create(true)){
            Sheet s = wb.createSheet("my sheet");
            Cell c = s.createRow(2).createCell(5);
            c.setCellValue("hello");
            
            CellStyle cs = wb.createCellStyle();
            Font f = wb.createFont();
            f.setColor(IndexedColors.WHITE.getIndex());
            cs.setFont(f);
            cs.setFillForegroundColor(IndexedColors.PINK.getIndex());
            cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            c.setCellStyle(cs);
            
            write(wb);
        }catch (IOException e) {
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
