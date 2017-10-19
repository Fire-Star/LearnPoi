package cn.domarvel.utils;

import cn.domarvel.exception.BaseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.InputStream;
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
    private static int PIXEL = 620;

    /**
     * 根据 Excel模板 创建 Excel 表格
     * 注意：标题必须为字符串或者其它，反正就是除了日期以外。
     * @param sheetModel 模板
     * @param outPutPath 表格创建成功后的输出路径
     * @param outPutFileName 输出文件名
     * @param insertData 插入数据
     * @param dataStartY 数据开始行
     * @param titleStartX 标题开始列的下标
     * @param titleStartY 标题开始行的下标
     * @throws Exception 异常
     */
    public static void createSheetByModel(InputStream sheetModel, String outPutPath, String outPutFileName, List<List<String>> insertData, int dataStartY,int titleStartX,int titleStartY) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook(sheetModel);
        Sheet sheet = wb.getSheetAt(0);//模板必须要有第一个Sheet。
        if(sheet == null){
            throw new BaseException(ErrorType.DANGER_TYPE,"模板不能为空！");
        }
        Row titleRow = sheet.getRow(titleStartY);//获得标题行
        if(titleRow == null){
            throw new BaseException(ErrorType.DANGER_TYPE,"模板不能没有标题行");
        }
        int tempCellCount = titleRow.getLastCellNum();
        int allCellCount = tempCellCount - titleStartX;

        //自动适应列宽
        for (int tempCell = 0; tempCell < allCellCount; tempCell++) {
            Cell cellObj = titleRow.getCell(tempCell+titleStartX);
            int cellStrLen = getMaxLengthInColumn(insertData,tempCell);//取出数据中字符长度的最大值。
            if(cellObj != null){
                String cellTitleStr = cellObj.toString();
                int len = cellTitleStr.length();
                cellStrLen = cellStrLen < len ? len : cellStrLen;
            }
            System.out.println(cellStrLen);
            sheet.setColumnWidth(tempCell + titleStartX,cellStrLen * PIXEL);
        }

        fillData(wb,sheet,insertData.size(),allCellCount,dataStartY,titleStartX,insertData);

        OutputStream outputStream = new FileOutputStream(outPutPath+outPutFileName);
        wb.write(outputStream);//输出
        sheetModel.close();
        outputStream.close();
        wb.close();
    }

    /**
     * 根据 Excel模板 创建 Excel 表格
     * 注意：标题必须为字符串或者其它，反正就是除了日期以外。
     * @param sheetModel 模板
     * @param outPutPath 表格创建成功后的输出路径
     * @param outPutFileName 输出文件名
     * @param insertData 插入数据
     * @param titleStartX 标题开始列的下标
     * @param titleStartY 标题开始行的下标
     * @throws Exception 异常
     */
    public static void createSheetByModel(InputStream sheetModel, String outPutPath, String outPutFileName, List<List<String>> insertData,int titleStartX,int titleStartY) throws Exception {
        createSheetByModel(sheetModel,outPutPath,outPutFileName,insertData,titleStartY+1,titleStartX,titleStartY);
    }

    /**
     *
     * 快速创建 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @param titleForegroundColor 标题前景色
     * @param titleTextColor 标题字体颜色
     * @param titleBorderColor 标题边框颜色
     * @throws Exception 异常
     */
    public static void createExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY , short titleForegroundColor,short titleTextColor,short titleBorderColor) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);//创建表


        CellStyle cellStyle = wb.createCellStyle();
        if( titleForegroundColor > -1 ){//如果有前景色就设置前景色，没有就是默认细黑色边框
            cellStyle.setFillForegroundColor(titleForegroundColor);//设置前景色
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        //设置标题细边框
        if(titleBorderColor > -1){//如果设置了边框颜色就设置边框颜色，否则就默认黑色边框
            cellStyle.setTopBorderColor(titleBorderColor);
            cellStyle.setBottomBorderColor(titleBorderColor);
            cellStyle.setLeftBorderColor(titleBorderColor);
            cellStyle.setRightBorderColor(titleBorderColor);
        }
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        if(titleTextColor > -1){
            //设置文字颜色
            Font font = wb.createFont();
            font.setColor(titleTextColor);
            cellStyle.setFont(font);
        }

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
            sheet.setColumnWidth(startIndexColumn,PIXEL * maxStrLen);
        }

        int allCellCount = titleName.length;//一共有多少列,列按照标题来
        int startRowIndex = startY+1;//内容的起始行
        int allRow = insertData.size();//内容一共有多少行

        fillData(wb,sheet,allRow,allCellCount,startRowIndex,startX,insertData);
        try {

            OutputStream outputStream = new FileOutputStream(outPutBathPath+outPutFileName);
            wb.write(outputStream);
            outputStream.close();
            wb.close();

        } catch (Exception e) {
            throw new BaseException(ErrorType.WARNING_TYPE,"生成的excel写入文件失败！");
        }
    }

    /**
     *  将内容填充到 Excel 中
     * @param wb 目标 Workbook
     * @param sheet Excel 表
     * @param allRow 数据一共有多少行
     * @param allCellCount 数据一共有多少列
     * @param startRowIndex 开始填充数据的行 index , index 从 0 开始
     * @param startX 开始填充数据的列
     * @param insertData 数据
     */
    public static void fillData(Workbook wb,Sheet sheet, int allRow, int allCellCount, int startRowIndex,int startX,List<List<String>> insertData){

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
    }

    /**
     *
     * 快速创建默认 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @param titleTextColor 标题文字色
     * @param titleBorderColor 标题边框色
     * @throws Exception 异常
     */
    public static void createDefaultStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY ,short titleTextColor, short titleBorderColor) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,startX,startY,(short)-1,titleTextColor,titleBorderColor);
    }
    /**
     *
     * 快速创建默认 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @throws Exception 异常
     */
    public static void createDefaultStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,startX,startY,(short)-1,(short)-1,(short)-1);
    }

    /**
     *
     * 快速创建默认 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @throws Exception 异常
     */
    public static void createDefaultStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[]) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,1,1,(short)-1,(short)-1,(short)-1);
    }

    /**
     *
     * 快速创建常用经典 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @param titleBorderColor 标题边框色
     * @throws Exception 异常
     */
    public static void createClassicStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY, short titleBorderColor) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,startX,startY,IndexedColors.LIGHT_BLUE.getIndex(), IndexedColors.WHITE.getIndex(),titleBorderColor);
    }
    /**
     *
     * 快速创建常用经典 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @param startX 整个表的起始位置X
     * @param startY 整个表的起始位置Y
     * @throws Exception 异常
     */
    public static void createClassicStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[], int startX, int startY) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,startX,startY,IndexedColors.LIGHT_BLUE.getIndex(), IndexedColors.WHITE.getIndex(),(short)-1);
    }

    /**
     *
     * 快速创建常用经典 Excel 表格
     * @param outPutBathPath 要类似于这样的路径：c:/base/path/
     * @param outPutFileName 随意
     * @param insertData 字符串List
     * @param sheetName 要创建的sheetName
     * @param titleName 标题名数组
     * @throws Exception 异常
     */
    public static void createClassicStyleExcel(String outPutBathPath, String outPutFileName, List<List<String>> insertData, String sheetName, String titleName[]) throws Exception {
        createExcel(outPutBathPath,outPutFileName,insertData,sheetName,titleName,1,1,IndexedColors.LIGHT_BLUE.getIndex(), IndexedColors.WHITE.getIndex(),(short)-1);
    }

    /**
     * 在excel内容数据中找到当前列的最长字符串长度
     * @param data excel内容数据
     * @param column 指定列
     * @return 当前列最长字符串长度
     */
    public static int getMaxLengthInColumn(List<List<String>> data ,int column){
        int max = 0;

        if(data == null){//如果没有数据就返回0
            return max;
        }

        for (List<String> tempData : data) {
            if(tempData == null){//如果没有当前行数据就遍历下一行
                continue;
            }
            if(tempData.size() <= column){
                continue;
            }
            String tempStr = tempData.get(column);
            if(tempStr == null){
                continue;
            }
            int tempLen = getExcelStrLen(tempStr);
            max = max < tempLen ? tempLen : max;
        }
        return max;
    }

    /**
     *  计算 Excel 中需要的长度，全是中文字符就返回中文字符串的长度，全是英文字符就返回英文字符长度的一半
     * @param tempStr 要被计算的字符串
     * @return 被计算后的长度
     */
    private static int getExcelStrLen(String tempStr) {
        int allCount = 0;
        String regex = "[\\x00-\\x7F]";
        int len = tempStr.length();
        int enCount = 0;
        for (int i = 0; i < len; i++) {
            String tempCharAt = String.valueOf(tempStr.charAt(i));
            if(tempCharAt.matches(regex)){
                enCount++;
            }
        }
        int chCount = len - enCount;
        allCount = enCount / 2 + chCount;
        return allCount;
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
