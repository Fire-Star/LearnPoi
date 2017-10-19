package cn.domarvel.poidemo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/18
 * Time: 20:05
 */
public class FillBackground {

    @Test
    public void fillBackground() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();

        XSSFSheet sheet = wb.createSheet("填充单元格背景颜色");
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);

        cell.setCellValue("Joy");
        sheet.setColumnWidth(1,520*3);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = wb.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());

        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);


        wb.write(new FileOutputStream("K:\\文件上传\\POI学习测试数据\\填充颜色.xlsx"));
        wb.close();
    }
}
