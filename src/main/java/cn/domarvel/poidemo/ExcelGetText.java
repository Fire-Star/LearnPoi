package cn.domarvel.poidemo;

import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/18
 * Time: 19:35
 */
public class ExcelGetText {
    @Test
    public void showExcelText() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook("K:\\文件上传\\POI学习测试数据\\技术人员名单.xlsx");

        XSSFExcelExtractor ex = new XSSFExcelExtractor(wb);
        ex.setIncludeSheetNames(false);
        System.out.println(ex.getText());
    }
}
