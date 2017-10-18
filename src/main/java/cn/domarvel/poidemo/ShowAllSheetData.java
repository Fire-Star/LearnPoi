package cn.domarvel.poidemo;

import cn.domarvel.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class ShowAllSheetData {


    //遍历xls文件
    public static void main(String[] args) throws IOException {

        XSSFWorkbook wb = new XSSFWorkbook("K:\\文件上传\\POI学习测试数据\\技术人员名单.xls");
        int sheetCount = wb.getNumberOfSheets();//包括隐藏的Sheet

        for (int i = 0; i < sheetCount; i++) {
            XSSFSheet tempSheet = wb.getSheetAt(i);

            if(tempSheet==null){//如果该excel文件中，一个Sheet都没有，那么就跳过。往往，目前这样获取时，这种情况不存在！！！
                continue;
            }
            int rowCount = tempSheet.getLastRowNum()+1;//行数量，PIO读取有可能会少一行，所以，我们+1
            System.out.println("#################################################");
            System.out.println("行数量："+rowCount);
            for (int row = 0; row < rowCount; row++) {
                XSSFRow tempRow = tempSheet.getRow(row);
                if(tempRow==null){//当该行没有定义的时候，就为null
                    System.out.println(row+"这是为null的行！");
                    continue;
                }
                int cellCount = tempRow.getLastCellNum(); // 我们在行的时候就出现了 少一行的情况，所以以后有需求时，这里少了一行，也不为怪，+1 就行了！
                for (int cell = 0; cell < cellCount; cell++) {
                    XSSFCell tempCell = tempRow.getCell(cell);
                    if(tempCell==null){//如果该单元格没有定义的时候，就为null
                        System.out.print("row="+row+" cell="+cell+" 这是为null的列！");
                        continue;
                    }
                    String value = ExcelUtils.getCellValue(tempCell);
                    System.out.print(value+":row="+row+" cell="+cell+"\t\t");
                }
                System.out.println();
            }
        }
        wb.close();
    }
}
