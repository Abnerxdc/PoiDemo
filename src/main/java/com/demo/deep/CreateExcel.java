package com.demo.deep;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.Date;

/**
 * Created by Administrator on 2017/6/13.
 */
public class CreateExcel {
    private static Logger logger = Logger.getLogger(CreateExcel.class);

    public static void create(){
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建Excel的工作sheet,对应到一个excel文档的tab
        HSSFSheet sheet = wb.createSheet("sheet1");

        //设置整列属性
        //sheet.setDefaultColumnStyle(short column, CellStyle style);

        // 设置excel每列宽度
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 3500);

        // 创建字体样式  字体、字体加粗、字体大小、字体颜色（下面一一对应）
        HSSFFont font = wb.createFont();
        font.setFontName("Verdana");
        font.setBoldweight((short) 1000);
        font.setFontHeight((short) 300);
        font.setColor(HSSFColor.BLUE.index);

        // 创建单元格样式
        HSSFCellStyle style = wb.createCellStyle();
        //设置单元格对齐方式 居中（横向）
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置单元格对齐方式 居中（纵向）
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //设置单元格背景填充色
        style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        // 设置边框  边框颜色、设置底部边框、设置左部边框、设置右部边框、设置上部边框
        style.setBottomBorderColor(HSSFColor.RED.index);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);

        style.setFont(font);// 设置字体

        // 创建当前sheet的第一行
        HSSFRow row = sheet.createRow(0);
        // 设定行的高度
        row.setHeight((short) 500);
        // 创建第一行第一个单元格
        HSSFCell cell = row.createCell(0);

        // 合并单元格(startRow，endRow，startColumn，endColumn)
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 2));

        // 给第一行第一个单元格使用上面的style 然后给单元格赋值
        cell.setCellStyle(style);
        cell.setCellValue("hello world");

        // 设置单元格内容 格式
        HSSFCellStyle style1 = wb.createCellStyle();
        style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));

        // 自动换行
        style1.setWrapText(true);

        row = sheet.createRow(2);

        // 创建第二行第一列的数据，应用style1 内容是时间
        cell = row.createCell(0);
        cell.setCellStyle(style1);
        cell.setCellValue(new Date());

        // 创建超链接
        HSSFHyperlink link = new HSSFHyperlink(HSSFHyperlink.LINK_URL);
        link.setAddress("http://www.baidu.com");
        cell = row.createCell(1);
        cell.setCellValue("百度");
        // 设定单元格的链接
        cell.setHyperlink(link);
        try{
            FileOutputStream os = new FileOutputStream("conf/create.xls");
            wb.write(os);
            os.close();
        }catch (Exception e){

        }

    }
}
