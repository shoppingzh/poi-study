package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

public class App {

    @Test
    public void merge() {
        try (Workbook wb = WorkbookFactory.create(true)) {
            Sheet s = wb.createSheet("my sheet");
            Row r0 = s.createRow(0);
            r0.createCell(0).setCellValue("基本信息");
            Row r2 = s.createRow(2);
            r2.createCell(0).setCellValue("姓名");
            r2.createCell(1).setCellValue("性别");
            r2.createCell(2).setCellValue("年龄");
            r2.createCell(3).setCellValue("联系电话");
            r2.createCell(4).setCellValue("地址");

            s.addMergedRegion(new CellRangeAddress(0, 1, 0, 4));
            // 将单元格居中
            CellStyle cs = wb.createCellStyle();
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            r0.getCell(0).setCellStyle(cs);

            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readMergeRegions() {
        try (Workbook wb = WorkbookFactory.create(new File("out/merge.xlsx"))) {
            Sheet s = wb.getSheetAt(wb.getActiveSheetIndex());
            for (int x = s.getFirstRowNum(); x <= s.getLastRowNum(); x++) {
                Row r = s.getRow(x);
                if (r == null) {
                    System.out.println("[empty row](index: " + x + ")");
                    continue;
                }
                for (int i = 0; i < r.getLastCellNum(); i++) {
                    Cell c = r.getCell(i);
                    if (c == null) {
                        System.out.print("[empty], ");
                        continue;
                    }
                    System.out.print(r.getCell(i).getStringCellValue() + ", ");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void unmerge() {
        try (Workbook wb = WorkbookFactory.create(new File("out/merge.xlsx"))) {
            Sheet s = wb.getSheetAt(wb.getActiveSheetIndex());

            // 处理合并
            for (int i = 0; i < s.getNumMergedRegions(); i++) {
                s.removeMergedRegion(i++);
            }

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
