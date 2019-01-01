package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void align() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet("mySheet");
            s.setDefaultColumnWidth(20);
            Row r0 = s.createRow(0);
            r0.setHeightInPoints(30f);
            Cell r0c0 = r0.createCell(0);
            r0c0.setCellValue("center text");
            CellStyle cs1 = wb.createCellStyle();
            cs1.setAlignment(HorizontalAlignment.CENTER);
            cs1.setVerticalAlignment(VerticalAlignment.CENTER);
            r0c0.setCellStyle(cs1);
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void horizontalAlign() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet("mySheet");
            Row r0 = s.createRow(0);
            setAlign(wb, r0.createCell(0), "general", HorizontalAlignment.GENERAL);
            setAlign(wb, r0.createCell(1), "left", HorizontalAlignment.LEFT);
            setAlign(wb, r0.createCell(2), "center", HorizontalAlignment.CENTER);
            setAlign(wb, r0.createCell(3), "right", HorizontalAlignment.RIGHT);
            setAlign(wb, r0.createCell(4), "fill", HorizontalAlignment.FILL);
            setAlign(wb, r0.createCell(5), "justify(自适应可自动换行)", HorizontalAlignment.JUSTIFY);
            setAlign(wb, r0.createCell(6), "center selection", HorizontalAlignment.CENTER_SELECTION);
            setAlign(wb, r0.createCell(7), "distrubuted", HorizontalAlignment.DISTRIBUTED);
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void verticalAlign() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet("mySheet");
            Row r0 = s.createRow(0);
            setAlign(wb, r0.createCell(0), "top", VerticalAlignment.TOP);
            setAlign(wb, r0.createCell(1), "fill", VerticalAlignment.CENTER);
            setAlign(wb, r0.createCell(2), "bottom", VerticalAlignment.BOTTOM);
            setAlign(wb, r0.createCell(3), "justify(自适应可自动换行)", VerticalAlignment.JUSTIFY);
            setAlign(wb, r0.createCell(4), "distrubuted", VerticalAlignment.DISTRIBUTED);
            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setAlign(Workbook wb, Cell cell, Object value, HorizontalAlignment align) {
        setValue(cell, value);
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(align);
        cell.setCellStyle(cs);
    }

    private static void setAlign(Workbook wb, Cell cell, Object value, VerticalAlignment align) {
        setValue(cell, value);
        CellStyle cs = wb.createCellStyle();
        cs.setVerticalAlignment(align);
        cell.setCellStyle(cs);
    }

    private static void setValue(Cell cell, Object value) {
        if (value instanceof Boolean) {
            cell.setCellValue((boolean) value);
        } else if (value instanceof Double) {
            cell.setCellValue((double) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
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
