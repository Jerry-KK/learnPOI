package cn.lethekk.learnapachepoi.service;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 导出模板需要分析
 * 分析下需要哪些操作
 *
 */
public class Test2 {


    //合并单元格


    //图片填充单元格

    //一个单元格连个标签


    public static void main(String[] args) {
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet newSheet1 = workbook.createSheet("newSheet1");

        HSSFRow row = newSheet1.createRow(0);

        HSSFCell cell = row.createCell(1);

        CellRangeAddress cellRangeAddress = new CellRangeAddress(1,1,2,4);
        newSheet1.addMergedRegion(cellRangeAddress);

        

    }
}
