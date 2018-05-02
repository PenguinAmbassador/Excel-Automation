/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

/**
 *
 * @author WoodmDav
 */
import java.io.File;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import tutorial.Employee;

public class WriteData {

    private static String[] columns = {"Name", "Email", "Date Of Birth", "Salary"};
    private static List<Employee> employees =  new ArrayList<>();

	// Initializing employees data to insert into the excel file
    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com", 
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com", 
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com", 
                dateOfBirth.getTime(), 1800000.0));
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // Obtain a workbook from the excel file
        Workbook workbook = WorkbookFactory.create(new File("src\\ConReport041918.xlsx"));

        //Establish old week and new week to prepare to copy headers
        Sheet lastWeekSheet = workbook.getSheetAt(4);
        //FIX replace with Date 
        String newSheetName = "Apr 16";
        workbook.createSheet(newSheetName);
        workbook.setSheetOrder("Apr 16", 4);
        Sheet newSheet = workbook.getSheetAt(4);
        
        //Copy headers from an old week
        Row header = lastWeekSheet.getRow(0);
        
        
        // Get Row at index 0
        Row rowTarg = newSheet.getRow(0);
        Row rowDest = newSheet.getRow(0);
        for(Cell cell : rowTarg){
            
        }
        for(Cell cell : row){
            
        }
        Cell cell = row.getCell(2);

        // Create the cell if it doesn't exist
        if (cell == null)
            cell = row.createCell(2);

        // Update the cell's value
        cell.setCellType(CellType.STRING);
        cell.setCellValue("Updated Value");

        // Write the output to the file
        FileOutputStream fileOut = new FileOutputStream("ConReport041918.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }
    
}

