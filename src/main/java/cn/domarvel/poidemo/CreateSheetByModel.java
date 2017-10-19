package cn.domarvel.poidemo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/19
 * Time: 10:55
 */
public class CreateSheetByModel {

    @Test
    public void createSheetByModel() throws Exception {
        InputStream inputStream = new FileInputStream("K:\\文件上传\\模板.xlsx");
        Workbook wb = new XSSFWorkbook(inputStream);

        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.createRow(2);
        row.createCell(1).setCellValue("Joy");
        row.createCell(2).setCellValue("杨舒粤");

        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\根据模板修改.xlsx");
        wb.write(outputStream);
        inputStream.close();
        outputStream.close();
        wb.close();
    }
}
