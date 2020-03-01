package com.dlrh.poi.utils;




import com.dlrh.poi.entity.Student;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;


public class SxssfWorkWiter {
    public static SXSSFWorkbook getHSSFWorkbook(String sheetName, String cellName, Class clazz, List<Student> dataList) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Field[] fields = clazz.getDeclaredFields();

        // SXSSFWorkbook支持excel2010
        SXSSFWorkbook sw = new SXSSFWorkbook();
        // 表名
        SXSSFSheet sheet = sw.createSheet(sheetName);
        // 创建行 【第一行  从零开始】
        SXSSFRow row = sheet.createRow(0);
        // 创建列 【第一列  从零开始】
        SXSSFCell cell = row.createCell(0);

        //设置单元格内容
        cell.setCellValue(cellName);
        //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,fields.length));

        // 创建样式
        CellStyle cellStyle = sw.createCellStyle();

        // 设置单元格的横向和纵向对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 给指定列设置样式
        cell.setCellStyle(cellStyle);

        //设置第一行
        SXSSFRow row1 = sheet.createRow(1);
        for (int i = 0; i<fields.length;i++){
            Field field = fields[i];
            String fieldName = field.getName();
            System.out.println(fieldName+"2222222222");
            row1.createCell(i).setCellValue(fieldName);

        }

        //设置数据行0000
        for (int j = 0 ; j<dataList.size();j++){
           Student obj = dataList.get(j);
            SXSSFRow row2 = sheet.createRow(j + 2);
            for (int k = 0; k<fields.length;k++){
                Field field = fields[k];
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0,1).toUpperCase() + fieldName.substring(1);
                Method declaredMethod = clazz.getDeclaredMethod(getMethodName);
                Object v = declaredMethod.invoke(obj);

                row2.createCell(k).setCellValue(v == null ? "":v.toString());

            }
        }

        return sw;
    }

}
