/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelcreator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author default
 */
public class ExcelCreator {
    
    public static void excelReader() {
        String excelFilePath = "excelTest.xlsx";
        MyExcelMgr reader = new MyExcelMgr();
        try {
            List<MyExcel> listMyExcel = reader.readBooksFromExcelFile(excelFilePath);
            System.out.println(listMyExcel);
        }
        catch(IOException e) {
            e.printStackTrace();
        }
        
    }
    
    public static void excelWriter() {
        WriteExcelMgr excelWriter = new WriteExcelMgr();

        List<WriteTable> listBook = excelWriter.getListBook();
        String excelFilePath = "result.xls";
        
        try {
            excelWriter.writeExcel(listBook, excelFilePath);
        }
        catch(IOException e) {
            e.printStackTrace();
        }
    }
    
    private static Workbook getWorkbook(FileInputStream inputStream, String excelFilePath)
        throws IOException {
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }
    
    public static void ExcelFormulaUpdate() {
        String excelFilePath = "resultFormat.xls";
        
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = getWorkbook(inputStream, excelFilePath);
            Sheet sheet = workbook.getSheetAt(0);
            ExcelFormulaCreate(workbook, sheet, sheet.getLastRowNum(), excelFilePath);
//            sheet.createRow(12).createCell(4).setCellFormula("SUM(D3:D8) + SUM(D3:D8)"); //if the row/cell didnt exist must be created.
//            sheet.getRow(12).getCell(4).setCellFormula("SUM(D3:D8) + SUM(D3:D8)");
//         
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
//            workbook.write(outputStream);
//            workbook.close();
            outputStream.close();
        }
        catch(IOException e) {
            e.printStackTrace();
        }
    }
    
    private static void ExcelFormulaCreate(Workbook workbook, Sheet sheet, int row, String file) {
        
        Row rowTotal = sheet.createRow(row + 2);
        Cell cellTotalText = rowTotal.createCell(2);
        cellTotalText.setCellValue("Total:");
         
        Cell cellTotal = rowTotal.createCell(3);
        cellTotal.setCellFormula("SUM(D3:D8)");
         
        try {
            
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            outputStream.close();
            
        }
        catch(IOException e) {
            
            e.printStackTrace();
            
        }
    } 
   public static void main(String[] args) throws IOException {
        ExcelFormulaUpdate();
    }
    
}
