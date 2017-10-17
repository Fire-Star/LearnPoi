package cn.domarvel.utils;

import cn.domarvel.exception.BaseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/17
 * Time: 12:45
 * Excel工具类：所有方法都经过深思熟虑实现，扩展强，功能实现具体，BUG 少，结构堪称完美！！！
 *
 * 都是为了让你少写代码，不！少加班！
 */
public class ExcelUtils {
    /**
     *
     * @param bathPath 要类似于这样的路径：c:/base/path/
     * @param fileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @throws IOException
     */
    public static void createExcel(String bathPath, String fileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);//创建表


        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);//设置标题细黑色边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//让文本居中

        Row titleRow = sheet.createRow(startY);//创建标题行
        //设置标题行内容
        for (int i = 0; i < titleName.length; i++) {
            String tempTitle = titleName[i];
            Cell tempCell = titleRow.createCell(i+startX);
            tempCell.setCellValue(tempTitle);
            tempCell.setCellStyle(cellStyle);
        }

        int allCellCount = titleName.length;//一共有多少列,列按照标题来
        int startRowIndex = startY+1;//内容的起始行
        int allRow = insertData.size();//内容一共有多少行

        for (int row = 0; row < allRow; row++) {//遍历所有的内容
            Row tempRow = sheet.createRow(row+startRowIndex);
            List<String> tempData = insertData.get(row);
            int nowCellCount = tempData.size();//实际的列数
            for (int cell = 0; cell < allCellCount; cell++) {

                if(nowCellCount<=cell){//如果当前遍历的列数大于实际的列数就退出当前行，设置下一行！
                    break;
                }
                Cell targetCell = tempRow.createCell(cell+startX);//创建列，设置列。
                targetCell.setCellValue(tempData.get(cell));
            }
        }
        try {

            OutputStream outputStream = new FileOutputStream(bathPath+fileName);
            wb.write(outputStream);
            outputStream.close();
            wb.close();

        } catch (Exception e) {
            throw new BaseException(StaticString.WARNING_TYPE,"生成的excel写入文件失败！");
        }
    }
}
