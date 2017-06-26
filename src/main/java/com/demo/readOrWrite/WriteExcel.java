package com.demo.readOrWrite;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;

/**
 * Created by Administrator on 2017/6/12.
 */
public class WriteExcel {
    private static JSONArray array = new JSONArray();
    Logger logger = Logger.getLogger(WriteExcel.class);

    public static void setArray(  ){
        JSONObject obj = new JSONObject();
        for(int i=0;i<2 ;i++){
            obj.put("KeyName","123"+i+99);
            obj.put("ColumnName","qwe"+i+99);
            obj.put("TableName","qwe"+i+99);
            obj.put("city","qwe"+i+99);
            array.add(obj);
        }
    }

    public static void getResult(){
        setArray();
        File file = new File("conf/test2.xls");
            setData(file,2);
    }

    public static void setData(File file, int ignoreRows) {

        //创建文件输出流
        OutputStream out = null;
        BufferedInputStream in = null;
        int rowSize = 0;
        try {
            in = new BufferedInputStream(new FileInputStream(
                    file));
            // 打开HSSFWorkbook
            POIFSFileSystem fs = new POIFSFileSystem(in);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFCell cell = null;
            for (int i = 0; i < array.size(); i++) {
                JSONObject jsonObject = array.getJSONObject(i);
                //只写sheet的第一个sheet
                HSSFSheet st = wb.getSheetAt(0);
                //从第五行开始写
                Row row = st.createRow(i + 5);
                //第几列
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(jsonObject.get("KeyName").toString());
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(jsonObject.get("ColumnName").toString());
                Cell cell2 = row.createCell(2);
                cell2.setCellValue(jsonObject.get("TableName").toString());
                Cell cell3 = row.createCell(3);
                cell3.setCellValue(jsonObject.get("city").toString());
            }

            in.close();
            out = new FileOutputStream(file);
            wb.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
