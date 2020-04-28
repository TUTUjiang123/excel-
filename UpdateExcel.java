package com.tutu.springboot;

import java.io.*;
import java.util.HashMap;

import jxl.*;
import jxl.format.*;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * Created by
 * hpt
 * on 2020/4/27.
 * 重庆渝欧跨境电子商务股份有限公司
 *  * Excel对比工具
 *  */
public class UpdateExcel {
    public static void main(String args[]) {
        WritableWorkbook book = null;
        HashMap<String, String> map = new HashMap<String, String>();
        try {
            // Excel获得文件
            Workbook wb = Workbook.getWorkbook(new File("D:\\Study\\exceltest\\cytest.xls"));
            // 打开一个文件的副本，并且指定数据写回到原文件
            book = Workbook.createWorkbook(new File("D:\\Study\\exceltest\\cytestX.xls"), wb);
            for(int s = 0; s<3; s++){
                Sheet sheet = book.getSheet(s);
                WritableSheet wsheet = book.getSheet(s);
                int colunms = sheet.getColumns();
                Boolean kg = false;
                for (int i = 0; i < sheet.getRows(); i++) {
                    if("subject_code".equals(sheet.getCell(1, i).getContents().trim())){
                        kg = true;
                    }
                    if("".equals(sheet.getCell(0, i+2).getContents().trim())){
                        kg = false;
                    }
                    if(kg){
                        Cell[] column1 = sheet.getRow(i+1);
                        Cell[] column2 = sheet.getRow(i+2);
                        for(int j = 2; j < colunms-1; j++){
                            if(!column1[j].getContents().trim().equals(column2[j].getContents().trim())){
                                Label label1 = new Label(j, i+1,
                                        column1[j].getContents().trim(),getDataCellFormat());
                                Label label2 = new Label(j, i+2,
                                        column2[j].getContents().trim(),getDataCellFormat());
                                wsheet.addCell(label1);
                                wsheet.addCell(label2);
                            }
                        }
                        i++;
                    }
                }
            }
            book.write();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                book.close();
                System.out.println("执行结束");
            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        }
    }

    // 设置标注的格式为黄底红字
    public static WritableCellFormat getDataCellFormat() {
        WritableCellFormat wcf = null;
        try {
            WritableFont wf = new WritableFont(WritableFont.createFont("宋体"), 11,
                    WritableFont.NO_BOLD, false);
            // 字体颜色
            wf.setColour(Colour.RED);
            wcf = new WritableCellFormat(wf);
            // 对齐方式
            wcf.setAlignment(Alignment.LEFT);
            wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
            // 设置上边框
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            // 设置下边框
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
            // 设置左边框
            wcf.setBorder(Border.LEFT, BorderLineStyle.THIN);
            // 设置右边框
            wcf.setBorder(Border.RIGHT, BorderLineStyle.THIN);
            // 设置背景色
            wcf.setBackground(Colour.YELLOW);
            // 自动换行
            wcf.setWrap(false);
        } catch (WriteException e) {
            e.printStackTrace();
        }
        return wcf;
    }


}
