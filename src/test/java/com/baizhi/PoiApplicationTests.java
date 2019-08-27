package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作簿==创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");
        //通过工作表创建行
        HSSFRow row = sheet.createRow(0);
        //通过行创建单元格
        HSSFCell cell = row.createCell(0);
        //给单元格赋值
        cell.setCellValue("第一个单元格");
        //把文件导出
        try {
            workbook.write(new FileOutputStream(new File("D:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test1() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("测试");
        //设置单元格宽度
        sheet.setColumnWidth(2, 15 * 256);
        //设置日期格式
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat("yyyy年MM月dd日");
        //把日期格式交给样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format);
        //创建标题行
        HSSFRow row = sheet.createRow(0);
        String[] strings = {"id", "姓名", "生日"};
        for (int i = 0; i < strings.length; i++) {
            row.createCell(i).setCellValue(strings[i]);
        }
        //填充内容
        User user1 = new User("1", "派大星", new Date());
        User user2 = new User("2", "海绵宝宝", new Date());
        User user3 = new User("3", "章鱼哥", new Date());
        List<User> list = new ArrayList<>();
        list.add(user1);
        list.add(user2);
        list.add(user3);
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            row1.createCell(0).setCellValue(list.get(i).getId());
            row1.createCell(1).setCellValue(list.get(i).getName());
            HSSFCell cell = row1.createCell(2);
            //将样式对象设置到当前单元格中
            cell.setCellStyle(cellStyle);
            cell.setCellValue(list.get(i).getBir());
        }
        //把文件导出
        try {
            workbook.write(new FileOutputStream(new File("D:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
