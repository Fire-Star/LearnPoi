package cn.domarvel.utils;

import cn.domarvel.exception.BaseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.List;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/17
 * Time: 12:45
 * Excel工具类：所有方法都经过深思熟虑实现，扩展强，功能实现具体，BUG 少，结构堪称完美！！！
 *
 * 都是为了让你写出优雅的代码，不！是少加班！
 */
public class ExcelUtils {
    private static SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
    /**
     *
     * @param bathPath 要类似于这样的路径：c:/base/path/
     * @param fileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @throws Exception 异常
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

            if(tempTitle == null){
                continue;
            }
            int startIndexColumn = i+startX;//计算开始插入列
            Cell tempCell = titleRow.createCell(startIndexColumn);
            tempCell.setCellValue(tempTitle);
            tempCell.setCellStyle(cellStyle);
            //计算列宽并且设置列宽，按照最大值设置
            int maxStrLen = getMaxLengthInColumn(insertData,i);
            maxStrLen = maxStrLen < tempTitle.length() ? tempTitle.length() : maxStrLen;
            System.out.println(maxStrLen);
            sheet.setColumnWidth(startIndexColumn,520*maxStrLen);
        }

        int allCellCount = titleName.length;//一共有多少列,列按照标题来
        int startRowIndex = startY+1;//内容的起始行
        int allRow = insertData.size();//内容一共有多少行

        CellStyle tempCellStyle = wb.createCellStyle();
        tempCellStyle.setAlignment(HorizontalAlignment.CENTER);//让文本居中

        for (int row = 0; row < allRow; row++) {//遍历所有的内容
            Row tempRow = sheet.createRow(row+startRowIndex);
            List<String> tempData = insertData.get(row);
            int nowCellCount = tempData.size();//实际的列数
            for (int cell = 0; cell < allCellCount; cell++) {

                if(nowCellCount<=cell){//如果当前遍历的列数大于实际的列数就退出当前行，设置下一行！
                    break;
                }
                Cell targetCell = tempRow.createCell(cell+startX);//创建列，设置列。
                targetCell.setCellStyle(tempCellStyle);//让文本居中
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

    /**
     * 在excel内容数据中找到当前列的最长字符串长度
     * @param data excel内容数据
     * @param column 指定列
     * @return 当前列最长字符串长度
     */
    public static int getMaxLengthInColumn(List<List<String>> data ,int column){
        int max = 0;
        for (List<String> tempData : data) {
            if(tempData.size() <= column){
                continue;
            }
            String tempStr = tempData.get(column);
            if(tempStr == null){
                continue;
            }
            int tempLen = tempStr.length();
            max = max < tempLen ? tempLen : max;
        }
        return max;
    }

    /**
     * 如果Cell为时间时，就通过默认的时间格式获取值
     * @param cell 目标Cell
     * @return 返回目标Cell字符串
     */
    public static String getCellValue(Cell cell){
        return getCellValue(cell,format);
    }

    /**
     * 获取Cell的值，自动判断类型
     * @param cell 目标Cell
     * @return 返回字符串
     */
    public static String getCellValue(Cell cell,SimpleDateFormat format) {
        if(cell == null){//如果为空就返回Null
            return null;
        }
        String o = null;
        int cellType = cell.getCellType();//获取当前Cell的类型
        switch (cellType) {
            //这里有个技巧，就是只判断特殊的几种类型，其它的类型就直接通过default处理。
            case Cell.CELL_TYPE_ERROR:// 5 当该单元格数据 ERROR 的时候,(故障)
                break;
            case Cell.CELL_TYPE_BLANK:// 3 当该单元格没有数据的时候
                o = "";
                break;
            case Cell.CELL_TYPE_NUMERIC:// 0 当该单元格数据为数字的时候
                o = getValueOfNumericCell(cell,format);
                break;
            case Cell.CELL_TYPE_FORMULA:// 2 当该单元格数据为公式的时候
                try {
                    o = getValueOfNumericCell(cell,format);
                } catch (IllegalStateException e) {
                    o = cell.getRichStringCellValue().toString();
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            default:
                o = cell.getRichStringCellValue().toString();
        }
        return o;
    }

    /**
     * 获取日期类型Cell的字符串值，通过默认的日期Pattern
     * @param cell 目标Cell
     * @return 返回目标Cell字符串
     */
    private static String getValueOfNumericCell(Cell cell) {
        return getValueOfNumericCell(cell,format);
    }

    /**
     * 获取日期类型Cell的字符串值
     * @param cell 目标Cell
     * @param format 日期Pattern
     * @return 返回目标Cell字符串
     */
    private static String getValueOfNumericCell(Cell cell, SimpleDateFormat format) {
        Boolean isDate = DateUtil.isCellDateFormatted(cell);
        String o = null;
        if (isDate) {
            o = format.format(cell.getDateCellValue());
        } else {
            o = String.valueOf(cell.getNumericCellValue());
        }
        return o;
    }
}
