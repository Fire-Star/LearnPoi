package cn.domarvel.poidemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class SetSheetValue {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();

        Sheet sheet = wb.createSheet("想你想你");
        Row firstRow = sheet.createRow(0);

        firstRow.createCell(0).setCellValue(5);
        firstRow.createCell(1).setCellValue(2);
        firstRow.createCell(2).setCellValue(0);
        firstRow.createCell(3).setCellValue("哈哈哈哈哈哈哈哈哈哈哈哈~~~");

        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\Sheet设置值演示.xls");
        wb.write(outputStream);
        outputStream.close();
    }
}
