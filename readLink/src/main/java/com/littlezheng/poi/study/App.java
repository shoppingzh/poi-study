package com.littlezheng.poi.study;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void readLink() {
        try (Workbook wb = WorkbookFactory.create(new File("out/1.xls"))) {
            Sheet s = wb.getSheetAt(wb.getActiveSheetIndex());
            for (Row r : s) {
                for (Cell c : r) {
                    Hyperlink link = c.getHyperlink();
                    if (link != null) {
                        System.out.println("超链接：" + link.getAddress());
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readLinkXs() {
        try (Workbook wb = WorkbookFactory.create(new File("out/1.xlsx"))) {
            Sheet s = wb.getSheetAt(wb.getActiveSheetIndex());
            for (Row r : s) {
                for (Cell c : r) {
                    Hyperlink link = c.getHyperlink();
                    if (link != null) {
                        System.out.println("超链接：" + link.getAddress());
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static void openLink(String addr) throws IOException {
        URLConnection conn = new URL(addr).openConnection();
        InputStream in = conn.getInputStream();
        byte[] buf = new byte[1024];
        List<Byte> all = new ArrayList<Byte>();
        int len = 1024;
        while ((len = in.read(buf)) != -1) {
            for (int i = 0; i < len; i++) {
                all.add(buf[i]);
            }
        }
        byte[] allBytes = new byte[all.size()];
        int i = 0;
        for (byte b : all) {
            allBytes[i++] = b;
        }
        System.out.println(new String(allBytes, "UTF-8"));
    }

    public static Object getValue(Cell c) {
        Object o = null;
        switch (c.getCellType()) {
        case BLANK:
            o = null;
            break;
        case BOOLEAN:
            o = c.getBooleanCellValue();
            break;
        case STRING:
            o = c.getStringCellValue();
            break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(c)) {
                o = c.getDateCellValue();
            } else {
                DecimalFormat df = new DecimalFormat("#.##");
                o = c.getNumericCellValue();
                System.out.println(df.format(o));
            }
        default:
            break;
        }

        return o;
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
