package cn.domarvel.poidemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class FirstPoi {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\新创建的POI.xls");
        wb.write(outputStream);
        outputStream.close();
    }
}
