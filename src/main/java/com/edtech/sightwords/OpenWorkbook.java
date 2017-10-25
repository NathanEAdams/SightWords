/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.edtech.sightwords;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author Nathan
 */
public class OpenWorkbook {
    public static void main(String args[]) throws Exception {

        File file = new File("C:\\Users\\Nathan\\Downloads\\Sight Words.xlsx");

        //Get the workbook instance for XLSX file
        try (FileInputStream fIP = new FileInputStream(file)) {
            //Get the workbook instance for XLSX file
            XSSFWorkbook workbook = new XSSFWorkbook(fIP);
            
            if (file.isFile() && file.exists()) {
                System.out.println(
                        "openworkbook.xlsx file open successfully.");
            } else {
                System.out.println("Error when opening openworkbook.xlsx file.");
            }
            XSSFSheet spreadsheet = workbook.getSheetAt(0);
            for (Row row : spreadsheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(
                                    cell.getNumericCellValue() + " \t\t ");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(
                                    cell.getStringCellValue() + " \t\t ");
                            break;
                    }
                }
                System.out.println();
            }
        }
    }
}
