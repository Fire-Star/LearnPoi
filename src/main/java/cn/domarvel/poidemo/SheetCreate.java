package cn.domarvel.poidemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class SheetCreate {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        wb.createSheet("FirstSheet");
        wb.createSheet("想你了，菇凉");
        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\Sheet演示.xls");
        wb.write(outputStream);
        outputStream.close();
    }
}
