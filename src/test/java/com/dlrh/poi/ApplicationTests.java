package com.dlrh.poi;

import com.dlrh.poi.entity.Student;
import com.dlrh.poi.utils.SxssfWorkWiter;
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
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
class ApplicationTests {

    @Test
    void contextLoads() {
    }

    @Test
    public void testCreateSheet () throws Exception {


        List<Student> s = new ArrayList<>();
        s.add(new Student("1","李颖","大连西岗区","女","99",25));
        s.add(new Student("2","李小颖","大连甘井子区","男","95",23));
        s.add(new Student("3","李中颖","大连中山区","女","93",22));
        s.add(new Student("4","李大颖","大连开发区","男","92",21));
        s.add(new Student("5","李颖儿","大连沙河口区","女","91",18));


        SXSSFWorkbook hssfWorkbook = SxssfWorkWiter.getHSSFWorkbook("sheet", "李颖详细", Student.class, s);

        FileOutputStream fileOutputStream = new FileOutputStream("D:\\student1.xlsx");

        hssfWorkbook.write(fileOutputStream);


        fileOutputStream.flush();
        fileOutputStream.close();
        hssfWorkbook.close();

    }


}
