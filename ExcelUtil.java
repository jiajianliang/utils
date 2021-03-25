package com.gw.ai.utils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class ExcelUtil {

    String filepath = "E:\\gw_ai_project\\src\\main\\resources\\语音质检测试用例.xlsx";
    Workbook wb = getExcel();
    Sheet sheet = wb.getSheetAt(0);


    public static void main(String[] args) {

        ExcelUtil excel = new ExcelUtil();
        excel.getExcel();
        Cell cell = excel.getCell(0,0);
        System.out.println(cell);
    }

    public Workbook getExcel(){

        File file=new File(filepath);
        if(!file.exists()){
            System.out.println("文件不存在");
            wb=null;
        }
        else {
            String fileType=filepath.substring(filepath.lastIndexOf("."));//获得后缀名
            try {
                InputStream is = new FileInputStream(filepath);
                if(".xls".equals(fileType)){
                    wb = new HSSFWorkbook(is);
                }else if(".xlsx".equals(fileType)){
                    wb = new XSSFWorkbook(is);
                }else{
                    System.out.println("格式不正确");
                    wb=null;
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        return wb;
    }

    public int getRouNum(){
        return sheet.getLastRowNum();
    }

    public int getCellNum(int num){
        Row row = sheet.getRow(num);
        int cellNum = row.getLastCellNum();
        return cellNum;
    }

    public Cell getCell(int row, int col){
        Row rowData = sheet.getRow(row);
        Cell  cellData = rowData.getCell(col);
        return cellData;
    }

}
