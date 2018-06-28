/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelcreator;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author default
 */
public class WriteExcelMgr {
    
    public List<WriteTable> getListBook() {
        
        int SIZE = 6;
        
        WriteTable[] arr = new WriteTable[SIZE];
        
        arr[0] = new WriteTable("Head First Java", "Kathy Serria", 79, 45, 65);
        arr[1] = new WriteTable("Effective Java", "Joshua Bloch", 36, 55, 65);
        arr[2] = new WriteTable("Clean Code", "Robert Martin", 42, 12, 12);
        arr[3] = new WriteTable("Thinking in Java", "Bruce Eckel", 35, 15, 1);
        arr[4] = new WriteTable("Fifth Row", "Bruce Eckel", 353, 151, 12);
        arr[5] = new WriteTable("Fifth asdfRow", "Bruce Ecsdfkel", 353, 1512, 12);
        
        List<WriteTable> listBook = new ArrayList<WriteTable>();
        
        for(int i = 0; i < SIZE; i++) {
            listBook.add(arr[i]);
        }

        return listBook;
    }
    
    private void writeBook(WriteTable aTable, Row row) {
        Cell cell = row.createCell(1);
        cell.setCellValue(aTable.getA());

        cell = row.createCell(2);
        cell.setCellValue(aTable.getB());

        cell = row.createCell(3);
        cell.setCellValue(aTable.getC());
        
        cell = row.createCell(4);
        cell.setCellValue(aTable.getD());
        
        cell = row.createCell(5);
        cell.setCellValue(aTable.getE());
    }
    private void createHeaderRow(Sheet sheet) {

        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);

        Row row = sheet.createRow(1);
        Cell cellTitle = row.createCell(1);

        cellTitle.setCellStyle(cellStyle);
        cellTitle.setCellValue("Info");

        Cell cellAuthor = row.createCell(2);
        cellAuthor.setCellStyle(cellStyle);
        cellAuthor.setCellValue("Info");

        Cell cellPrice = row.createCell(3);
        cellPrice.setCellStyle(cellStyle);
        cellPrice.setCellValue("Element");
        
        Cell cellElement = row.createCell(4);
        cellElement.setCellStyle(cellStyle);
        cellElement.setCellValue("Element");
        
        Cell cellElement2 = row.createCell(5);
        cellElement2.setCellStyle(cellStyle);
        cellElement2.setCellValue("Element");
    }
    
    private Workbook getWorkbook(String excelFilePath)
            throws IOException {
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }
    
    public void writeExcel(List<WriteTable> listTable, String excelFilePath) throws IOException {
        Workbook workbook = getWorkbook(excelFilePath);
        Sheet sheet = workbook.createSheet();
        createHeaderRow(sheet);
        
        int rowCount = 2;

        for (WriteTable aTable : listTable) {
            Row row = sheet.createRow(++rowCount);
            writeBook(aTable, row);
        }

        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }
    }
}
