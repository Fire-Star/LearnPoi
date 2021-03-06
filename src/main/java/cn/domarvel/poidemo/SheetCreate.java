package cn.domarvel.poidemo;

import cn.domarvel.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.Test;

import java.io.*;
import java.util.LinkedList;
import java.util.List;

public class SheetCreate {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        wb.createSheet("FirstSheet");
        wb.createSheet("想你了，菇凉");
        OutputStream outputStream = new FileOutputStream("K:\\文件上传\\Sheet演示.xls");
        wb.write(outputStream);
        outputStream.close();
    }

    @Test
    public void createSheetByUtils() throws Exception {
        String []title = {"姓名","简介"};
        List<List<String>> insertData = new LinkedList<>();

        List<String> tempData01 = new LinkedList<>();
        tempData01.add("FireLang");
        tempData01.add("很久很久以前，少年对黑客痴迷！");
        insertData.add(tempData01);

        List<String> tempData02 = new LinkedList<>();
        tempData02.add("MoonFollow");
        tempData02.add("在路途中，少年遇到了 Joy 并且深深深深的爱上了她！");
        insertData.add(tempData02);

        ExcelUtils.createExcel("K:\\文件上传\\","简介.xlsx",insertData,"简介",title,1,1,(short)-1,(short)-1,(short)-1);
    }

    @Test
    public void createClassicStyleSheet() throws Exception {
        String []title = {"姓名","简介"};

        List<List<String>> insertData = new LinkedList<>();

        List<String> tempData01 = new LinkedList<>();
        tempData01.add("FireLang");
        tempData01.add("很久很久以前，少年对黑客痴迷！");
        insertData.add(tempData01);

        List<String> tempData02 = new LinkedList<>();
        tempData02.add("MoonFollow");
        tempData02.add("在路途中，少年遇到了 Joy 并且深深深深的爱上了她！");
        insertData.add(tempData02);

        ExcelUtils.createClassicStyleExcel("K:\\文件上传\\","常用经典样式.xlsx",insertData,"常用经典样式",title);
    }

    @Test
    public void createSheetByModelTest() throws Exception {
        List<List<String>> insertData = new LinkedList<>();

        List<String> tempData01 = new LinkedList<>();
        tempData01.add("MoonFollow");
        tempData01.add("胡艺宝");
        insertData.add(tempData01);

        List<String> tempData02 = new LinkedList<>();
        tempData02.add("Joy");
        tempData02.add("杨舒粤");
        insertData.add(tempData02);

        List<String> tempData03 = new LinkedList<>();
        tempData03.add("FireLang");
        tempData03.add("胡艺宝");
        insertData.add(tempData03);

        ExcelUtils.createSheetByModel(new FileInputStream("K:\\文件上传\\模板.xlsx"),"K:\\文件上传\\","自动根据模板创建.xlsx",insertData,1,1);
    }
}
