package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void shiftRows(){
        try(Workbook wb = WorkbookFactory.create(true)){
            Sheet s = wb.createSheet("my sheet");
            for(int i=0;i<1000;i++){
                Row r = s.createRow(i);
                Cell c = r.createCell(0);
                c.setCellValue(String.valueOf(i));
                c.setCellType(CellType.STRING);
            }
            
            s.shiftRows(5, 10, -5);
            
            write(wb);
        }catch(IOException e){
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
