package cn.domarvel.poidemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

public class CreateDateSheet {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();

        Sheet sheet = wb.createSheet("每一秒");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("都在想你！");
        row.createCell(1).setCellValue("比如，这一秒：");
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue("但是呢，这代表程序员的时间！我想换个更浪漫的时间，比如：");

        //第一种创建方式
        Cell firstCell = row.createCell(4);
        firstCell.setCellValue(new Date());

        //两种时间设置方式不同，样式相同
        CellStyle cellStyle = wb.createCellStyle();
        DataFormat dataFormat =  wb.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd HH:mm:ss"));

        //为单元格设置样式
        firstCell.setCellStyle(cellStyle);

        row.createCell(5).setCellValue("就算只有这一生，我也会数着秒去想你！比如这一秒：");

        //第二种创建方式
        Cell secondCell = row.createCell(6);
        secondCell.setCellValue(Calendar.getInstance());

        //设置单元格样式
        secondCell.setCellStyle(cellStyle);

        row.createCell(7).setCellValue("你看，一秒时间都没我想你想得快，那么这一生就是一万年！");

        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\Sheet数着秒去想你.xls");
        wb.write(outputStream);
        outputStream.close();
    }
}
