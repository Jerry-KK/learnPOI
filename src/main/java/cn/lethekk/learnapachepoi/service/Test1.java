package cn.lethekk.learnapachepoi.service;

import com.sun.xml.internal.messaging.saaj.util.ByteOutputStream;
import org.apache.poi.hssf.usermodel.*;

import java.io.*;

/**
 *
 *
 */
public class Test1 {
    public static void main(String[] args) throws IOException {



        String filePath = "e:/Demo/test/first.xls";
        String filePath2 = "e:/Demo/test/f2.xls";
        File file = new File(filePath);
       /* try {
            file.createNewFile();
            System.out.println(1);
        } catch (IOException e) {
            e.printStackTrace();
        }*/


        //HSSFWorkbook：工作簿，代表一个excel的整个文档
        //InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook();                           //创建一个工作簿
        //HSSFWorkbook workbook2 = new HSSFWorkbook(inputStream);             //创建一个关联输入流的工作簿

        HSSFSheet sheet1 = workbook.createSheet("sheet1");          //创建一个新的sheet，名称为"sheet1"，传入的名称不能为空，不能重复

        HSSFSheet findsheet = workbook.getSheet("sheet1");              //通过名称获取sheet
        HSSFSheet findsheet2 = workbook.getSheetAt(0);                  //通过索引获取sheet，索引从0开始
        int num = workbook.getNumberOfSheets();                               //获取sheet个数


       // HSSFCellStyle cellStyle = workbook.createCellStyle();//创建单元格格式,还没发现要怎么用

        HSSFRow row = sheet1.createRow(0);

        for (int i = 0; i < 5; i++) {
            row.createCell(i).setCellValue("title"+i);
            row.setHeightInPoints(100f);                                //设置行的高度
        }
        for (int i = 1; i < 6; i++) {
            HSSFRow rown = sheet1.createRow(i);
            for (int j = 0; j < 5; j++) {
                rown.createCell(j).setCellValue("value"+i+":"+j);
            }
        }
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);


    }
}
