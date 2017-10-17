package cn.domarvel.poidemo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ShowAllSheetData {
    //遍历xls文件
    public static void main(String[] args) throws IOException {
        InputStream inputStream = new FileInputStream("K:\\文件上传\\POI学习测试数据\\技术部人员.xls");
        POIFSFileSystem fs = new POIFSFileSystem(inputStream);
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        int sheetCount = wb.getNumberOfSheets();//包括隐藏的Sheet

        for (int i = 0; i < sheetCount; i++) {
            HSSFSheet tempSheet = wb.getSheetAt(i);

            if(tempSheet==null){//如果该excel文件中，一个Sheet都没有，那么就跳过。往往，目前这样获取时，这种情况不存在！！！
                continue;
            }
            int rowCount = tempSheet.getLastRowNum()+1;//行数量，PIO读取有可能会少一行，所以，我们+1
            System.out.println("#################################################");
            System.out.println("行数量："+rowCount);
            for (int row = 0; row < rowCount; row++) {
                HSSFRow tempRow = tempSheet.getRow(row);
                if(tempRow==null){//当该行没有定义的时候，就为null
                    System.out.println(row+"这是为null的行！");
                    continue;
                }
                int cellCount = tempRow.getLastCellNum(); // 我们在行的时候就出现了 少一行的情况，所以以后有需求时，这里少了一行，也不为怪，+1 就行了！
                for (int cell = 0; cell < cellCount; cell++) {
                    HSSFCell tempCell = tempRow.getCell(cell);
                    if(tempCell==null){//如果该单元格没有定义的时候，就为null
                        System.out.print("row="+row+" cell="+cell+" 这是为null的列！");
                        continue;
                    }
                    String value = getCellData(tempCell);
                    System.out.print(value+":row="+row+" cell="+cell+"\t\t");
                }
                System.out.println();
            }
        }
        inputStream.close();
    }

    public static String getCellData(HSSFCell targetCell){
        int type = targetCell.getCellType();
        String value = "";//如果没有获取到数据就为 空字符串
        switch (type){
            case HSSFCell.CELL_TYPE_NUMERIC: // 0 当该单元格数据为数字的时候
                value = String.valueOf(targetCell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_STRING: // 1 当该单元格数据为字符串的时候
                value = targetCell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_FORMULA: // 2 当该单元格数据为公式的时候
                value = targetCell.getCellFormula();
                break;
            case HSSFCell.CELL_TYPE_BLANK: // 3 当该单元格数据为空的时候
                value = "BLANK";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN: // 4 当该单元格数据为布尔值的时候
                value = String.valueOf(targetCell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_ERROR: // 5 当该单元格数据 ERROR 的时候,(故障)
                value = "ERROR";
                break;
        }

        return value;
    }
}
