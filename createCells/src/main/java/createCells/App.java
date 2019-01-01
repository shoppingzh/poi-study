package createCells;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class App {

    @Test
    public void createCell() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet("mySheet");
            Row r0 = s.createRow(0);
            Cell r0c0 = r0.createCell(0);
            r0c0.setCellValue("类型");
            Cell r0c1 = r0.createCell(1);
            r0c1.setCellValue("表现");
            CellStyle headerStyle = headerStyle(wb);
            for (Cell c : r0) {
                c.setCellStyle(headerStyle);
            }

            Row r1 = s.createRow(1);
            r1.createCell(0).setCellValue("数值");
            r1.createCell(1).setCellValue(1.25);

            Row r2 = s.createRow(2);
            r2.createCell(0).setCellValue("布尔");
            r2.createCell(1).setCellValue(false);

            Row r3 = s.createRow(3);
            r3.createCell(0).setCellValue("字符串");
            r3.createCell(1).setCellValue("Hello, world!");

            Row r4 = s.createRow(4);
            r4.createCell(0).setCellValue("Calendar");
            r4.createCell(1).setCellValue(Calendar.getInstance());

            Row r5 = s.createRow(5);
            r5.createCell(0).setCellValue("Date");
            r5.createCell(1).setCellValue(Calendar.getInstance().getTime());

            write(wb);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private CellStyle headerStyle(Workbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setColor(Font.COLOR_RED);
        cs.setFont(f);
        return cs;
    }

    @Test
    public void createDateCell() {
        try (Workbook wb = WorkbookFactory.create(false)) {
            Sheet s = wb.createSheet("mySheet");
            Row r1 = s.createRow(1);
            Cell r1c0 = r1.createCell(0);
            r1c0.setCellValue(new Date());

            CellStyle cs = wb.createCellStyle();
            cs.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy h:mm"));
            r1c0.setCellStyle(cs);
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
