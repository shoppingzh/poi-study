package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void borders(){
        try(Workbook wb = WorkbookFactory.create(true)){
            Sheet s = wb.createSheet("my sheet");
            Row r = s.createRow(2);
            Cell c = r.createCell(5);
            r.setHeightInPoints(50f);
            c.setCellValue("hello");
            
            CellStyle cs = wb.createCellStyle();
            c.setCellStyle(cs);
            cs.setBorderTop(BorderStyle.THIN);
            cs.setTopBorderColor(IndexedColors.RED.getIndex());
            cs.setBorderRight(BorderStyle.MEDIUM_DASHED);
            cs.setRightBorderColor(IndexedColors.BLUE.getIndex());
            cs.setBorderBottom(BorderStyle.DOTTED);
            cs.setBottomBorderColor(IndexedColors.PINK.getIndex());
            cs.setBorderLeft(BorderStyle.DOUBLE);
            cs.setLeftBorderColor(IndexedColors.YELLOW.getIndex());
            
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
