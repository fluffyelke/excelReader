/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelcreator;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
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
public class MyExcelMgr {
    private Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();

            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();

            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            }

        return null;
    }
    
    private Workbook getWorkbook(FileInputStream inputStream, String excelFilePath)
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
    
    public List<MyExcel> readBooksFromExcelFile(String excelFilePath) throws IOException {
    List<MyExcel> listMyExcel = new ArrayList<>();
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
 
    Workbook workbook = getWorkbook(inputStream, excelFilePath);
    Sheet firstSheet = workbook.getSheetAt(0);
    Iterator<Row> iterator = firstSheet.iterator();
 
    while (iterator.hasNext()) {
        Row nextRow = iterator.next();
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        MyExcel aExcel = new MyExcel();
 
        while (cellIterator.hasNext()) {
            Cell nextCell = cellIterator.next();
            int columnIndex = nextCell.getColumnIndex();
            
            switch (columnIndex) {
            case 0:
                aExcel.setA((String) getCellValue(nextCell));
                break;
            case 1:
                aExcel.setB((String) getCellValue(nextCell));
                break;
            case 2:
            {
                double d = (double)getCellValue(nextCell);
                int myColumn = (int)d;
                aExcel.setC(myColumn);
            }
                break;
            case 3:
                {
                    double d = (double)getCellValue(nextCell);
                    int myColumn = (int)d;
                    aExcel.setD(myColumn);
                }
                break;
            case 4:
                {
                    double d = (double)getCellValue(nextCell);
                    int myColumn = (int)d;
                    aExcel.setE(myColumn);
                }
                break;
            }
            
          
 
 
        }
        listMyExcel.add(aExcel);
    }
 
//    workbook.close();
    inputStream.close();
 
    return listMyExcel;
}
}
