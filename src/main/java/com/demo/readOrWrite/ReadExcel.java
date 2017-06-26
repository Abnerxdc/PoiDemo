package com.demo.readOrWrite;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Administrator on 2017/6/12.
 */
public class ReadExcel {
    private static Logger logger = Logger.getLogger(ReadExcel.class);
    public static void getResult(){
        File file = new File("conf/create.xls");
        try{
            getDate(file,0);
        }catch (Exception e){
            logger.error(" io 输入流异常 ： "+e);
        }

    }

    /**
     * 读取Excel的内容
     * @param file 读取数据的源Excel
     * @param ignoreRows 读取数据忽略的行数，比喻行头不需要读入 忽略的行数为1
     * @return 读出的Excel中数据的内容
     * @throws FileNotFoundException
     * @throws IOException
     */
    public static void getDate(File file, int ignoreRows)throws FileNotFoundException, IOException{
        List<String[]> result = new ArrayList<String[]>();
        //创建文件输入流
        BufferedInputStream in  = new BufferedInputStream(new FileInputStream(
                file));
        // 打开HSSFWorkbook
        POIFSFileSystem fs = new POIFSFileSystem(in);
        //HSSFWorkbook提供读写.xls格式档案的功能。 其他类型的excel使用相应的workbook打开，例如.xlsx格式的Excel文档->使用XSSFWorkbook打开
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        //关于Workbook的创建，上面两句话可以用一句：Workbook wb =  WorkbookFactory.create(in);
        //对于加密的excel还要使用Biff8EncryptionKey.setCurrentUserPassword(*******);

        HSSFCell cell = null;
        //如果有多个工作页，循环打开
        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
            //打开这个工作页
            HSSFSheet st = wb.getSheetAt(sheetIndex);
            //设置循环 一行行 往外取值 一般前几行都是介绍，这里忽略掉给出的忽略行数
            for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {
                //获取当前行内容
                HSSFRow row = st.getRow(rowIndex);
                //先做一个放空处理
                if (row == null) {
                    continue;
                }
                //开始循环一个个取值
                for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
                    cell = row.getCell(columnIndex);
                    //和数组一样都是从0开始取的，虽然上面的溢出了一点点，这里补个判断
                    if(cell!=null){
                        //这里都按照String类型取 可以使用 switch（cell.getCellType()） case HSSFCell.CELL_TYPE_STRING: case HSSFCell.CELL_TYPE_NUMERIC: 直接取成自己想要的类型
                        // 数字有数字的取值方式，字符串有字符串的取值方式String value = cell.getStringCellValue();
                        // 上面这种比较麻烦，这里直接用下面获取内容，虽然不推荐。但是比较方便....
                        String value = cell.toString();
                        logger.info("第 "+rowIndex+" 行 ， 第 "+columnIndex+" 列的值为： "+value);
                    }
                }
                logger.info("共计："+st.getLastRowNum()+" 行  "+row.getLastCellNum()+" 列");
            }
        }
    }

}
