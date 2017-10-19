package cn.domarvel.poidemo;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/19
 * Time: 10:36
 */
public class CellMerge {

    @Test//单元格合并
    public void cellMerge() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("合并单元格");

        Row row = sheet.createRow(1);

        row.createCell(1).setCellValue("L*ve Joy");
        sheet.addMergedRegion(new CellRangeAddress(1,1,1,2));
        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\合并单元格.xlsx");
        wb.write(outputStream);
        wb.close();
    }
}
