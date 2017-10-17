package cn.domarvel.utils;

import java.lang.reflect.Field;
import java.util.LinkedList;
import java.util.List;

/**
 * Create by MoonFollow (or named FireLang)
 * Only For You , Joy
 * Date: 2017/10/17
 * Time: 12:57
 * 属性工具类，所有方法都经过深思熟虑实现，扩展强，功能实现具体，BUG 少，结构堪称完美！！！
 *
 * 都是为了让你少写代码，不！少加班！
 */
public class FieldUtils {
    /**
     * 在指定Class对象里面获取指定的属性
     * @param fieldName 属性名
     * @param targetClazz Class对象
     * @return
     */
    public static Field getTargetField(String fieldName, Class targetClazz){
        Field result = null;
        while (targetClazz != null){
            try {
                result = targetClazz.getDeclaredField(fieldName);
                return result;
            } catch (NoSuchFieldException e) {
                targetClazz = targetClazz.getSuperclass();
            }
        }
        return result;
    }

    /**
     * 将 List<Object> 中的属性按照属性数组顺序提取城 List<List<String>>，注意：使用该方法的前提是，你要提取的属性值都是字符串！
     * @param insertData 被转换的数据源
     * @param objectProName 数据属性
     * @return
     */
    public static List<List<String>> objectListProToStrList(List<?> insertData, String [] objectProName){
        List<List<String>> result = new LinkedList<>();
        for (Object tempInsert : insertData) {
            List<String> tempDate = objectProToStrList(tempInsert,objectProName);
            result.add(tempDate);
        }
        return result;
    }

    /**
     * 提取指定对象里面的指定属性，并且按照指定属相数组的顺序提取。注意：使用该方法的前提是，你要提取的属性值都是字符串！
     * @param data 指定对象
     * @param objectProName 指定属性字符串数组
     * @return 指定提取的属性List
     */
    public static List<String> objectProToStrList(Object data , String []objectProName){
        List<String> tempItem = new LinkedList<>();//创建装指定属性值的容器
        Class tempClazz = data.getClass();//获得该对象的Class对象
        for (String tempPro : objectProName) {
            try {
                Field tempField = FieldUtils.getTargetField(tempPro,tempClazz);
                tempField.setAccessible(true);//设置当前属性可访问
                String tempProValue = (String) tempField.get(data);
                if(tempProValue == null){//如果当前属性为 Null 那么就输出空字符串。
                    tempProValue = "";
                }
                tempItem.add(tempProValue);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return tempItem;
    }
}
