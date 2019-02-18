package com.littlezheng.poi.study;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadUtils {

    public static List<String[]> read(File file, int sheetIndex) {
        List<String[]> rows = new ArrayList<String[]>();
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet s = wb.getSheetAt(sheetIndex);
            for (Row r : s) {
                short maxCol = r.getLastCellNum();
                String[] row = new String[maxCol];
                for (int i = 0; i < maxCol; i++) {
                    row[i] = getCellValue(r.getCell(i));
                }
                rows.add(row);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return rows;
    }

    private static String getCellValue(Cell c) {
        String value = null;
        switch (c.getCellType()) {
        case STRING:
            value = c.getStringCellValue();
            break;
        case BOOLEAN:
            value = String.valueOf(c.getBooleanCellValue());
            break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(c)) {
                value = new SimpleDateFormat("yyyy-MM-dd").format(c.getDateCellValue());
            } else {
                DecimalFormat df = new DecimalFormat("#.##############");
                value = df.format(c.getNumericCellValue());
            }
            break;
        case BLANK:
        default:
            value = null;
            break;
        }
        return value;
    }

    public static void main(String[] args) {
        List<String[]> rows = read(new File("d:/1.xlsx"), 2);
        for (String[] row : rows) {
            System.out.println(Arrays.toString(row));
        }
    }

}
