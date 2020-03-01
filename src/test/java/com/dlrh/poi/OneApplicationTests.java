package com.dlrh.poi;

import com.dlrh.poi.entity.Student;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import org.springframework.boot.test.context.SpringBootTest;



import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
class OneApplicationTests {


    @Test
    void contextLoads() {
    }

    @Test
    public void testCreateSheet () throws Exception {

        System.out.println("ldkfasdjflksadjfksa");

        List<Student> s = new ArrayList<>();
        s.add(new Student("1","李颖","大连西岗区","女","99",25));
        s.add(new Student("2","李小颖","大连甘井子区","男","95",23));
        s.add(new Student("3","李中颖","大连中山区","女","93",22));
        s.add(new Student("4","李大颖","大连开发区","男","92",21));
        s.add(new Student("5","李颖儿","大连沙河口区","女","91",18));



        // 【SXSSFWorkbook----------->sheet----------->row------------>cell】


        Class<Student> studentClass = Student.class;
        Field[] fields = studentClass.getDeclaredFields();

        // SXSSFWorkbook支持excel2010
        SXSSFWorkbook sw = new SXSSFWorkbook();
        // 表名
        SXSSFSheet sheet = sw.createSheet("学生信息表");
        // 创建行 【第一行  从零开始】
        SXSSFRow row = sheet.createRow(0);
        // 创建列 【第一列  从零开始】
        SXSSFCell cell = row.createCell(0);

        //设置单元格内容
        cell.setCellValue("学员考试成绩一览表");
        //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));

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

        //设置数据行
        for (int j = 0 ; j<s.size();j++){
            Student obj = s.get(j);
            SXSSFRow row2 = sheet.createRow(j + 2);
            for (int k = 0; k<fields.length;k++){
                Field field = fields[k];
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0,1).toUpperCase() + fieldName.substring(1);
                Method declaredMethod = studentClass.getDeclaredMethod(getMethodName);
                Object v = declaredMethod.invoke(obj);

                row2.createCell(k).setCellValue(v == null ? "":v.toString());

            }
        }

        // 新建导出流
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\student.xlsx");
        // 将文件写出
        sw.write(fileOutputStream);

        fileOutputStream.flush();

        fileOutputStream.close();

        sw.close();
    }

}
