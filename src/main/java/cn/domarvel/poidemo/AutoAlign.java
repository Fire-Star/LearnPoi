package cn.domarvel.poidemo;
import cn.domarvel.utils.StaticString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.FileOutputStream;
import java.io.IOException;

public class AutoAlign {

    @Test
    public void test01() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("FirstSheet");
        sheet.autoSizeColumn((short)0,true);
        sheet.autoSizeColumn((short)1,true);

        /*sheet.setColumnWidth(0,"1111111111111111111111111111111".length()*512);*/
        XSSFRow firstRow = sheet.createRow(0);
        XSSFCell cell0 = firstRow.createCell(0);
        cell0.setCellValue("1111111111111111111111111111111");

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell0.setCellStyle(cellStyle);
        wb.write(new FileOutputStream(StaticString.DEFAULTPATH+"自动单元格对齐.xlsx"));
        wb.close();
    }

    @Test
    public void regexToGetWidth(){
        String regex = "[\\x00-\\x7F]";
        String sourceStr = "123abcABC,:;胡艺宝";
        int len = sourceStr.length();
        int countEn = 0;
        for (int i = 0; i < len; i++) {
            String itemStr = sourceStr.charAt(i)+"";
            if(itemStr.matches(regex)){
                countEn++;
            }
        }
    }

    public static void setColumnWidthAuto(Sheet sheet,int columnIndex){
        int allRowCount = sheet.getLastRowNum()+1;
        for (int rowCount = 0; rowCount < allRowCount; rowCount++) {
            Row tempRow = sheet.getRow(rowCount);
            if(tempRow == null){
                continue;
            }
            Cell targetCell = tempRow.getCell(columnIndex);
            if(targetCell == null){
                continue;
            }
        }
    }

}
