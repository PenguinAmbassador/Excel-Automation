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
import java.util.LinkedList;
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
        // Obtain a Connectivity Report from xlsx
        //This report should be a week older than the weekly report
        Workbook connReport = WorkbookFactory.create(new File("src\\ConReport041618.xlsx"));

        //Obtain Weekly Report. Used to build Connectivity Report.
        Workbook weeklyReport = WorkbookFactory.create(new File("src\\WeeklyReport041618.xlsx"));
        
        //Establish old week and new week to prepare to copy headers
        String newSheetName = "Apr 16";//FIX replace with Date 
        connReport.createSheet(newSheetName);
        connReport.setSheetOrder("Apr 16", 4);//move new sheet in front of other weeks
        Sheet newSheet = connReport.getSheetAt(4);//grab the new sheet
        Sheet lastWeekSheet = connReport.getSheetAt(5);//grab previous week
        Sheet weeklySheet = weeklyReport.getSheetAt(0);//Grab the only sheet in Weekly Report
        
        // STEP 1: Copy headers from an old week 
        Row rowSrc = lastWeekSheet.getRow(0);
        Row rowDest = newSheet.createRow(0);
        int colNum = 0;//helps iterate through the columns
        for(Cell cell : rowSrc){
            Cell currentCell = rowDest.createCell(colNum);
            currentCell.setCellValue(cell.getStringCellValue());
            currentCell.setCellStyle(cell.getCellStyle());
            newSheet.autoSizeColumn(colNum);//FIX: Not resizing columns...
            colNum++;
        }
        
        //STEP 2: Copy data from Weekly Report
        for(int i = 1; i < weeklySheet.getPhysicalNumberOfRows(); i++){ //copy every row but row 1
            //System.out.println("i: " + i);
            Row srcRow = weeklySheet.getRow(i); //Track source row in Weekly Report
            Row targRow = newSheet.createRow(i); //Track target row in newSheet in the Connectivity report
            for(int j = 0; j < 12; j++){ //copy every cell
                //System.out.println("j: " + j);
                //System.out.println(srcRow.getCell(j));
                Cell srcCell = srcRow.getCell(j); //track source cell in Weekly Report
                if(srcCell!=null){
                    if(j == 0){
                        Cell currentCell = targRow.createCell(j);
                        currentCell.setCellValue(i); //store row as #
                    }
                    if(j > 0 && j < 3){ //skip col 0 and 3 which are ID and Email
                        Cell currentCell = targRow.createCell(j);
                        currentCell.setCellValue(srcCell.getStringCellValue());
                        newSheet.autoSizeColumn(j);//FIX: Not resizing columns...
                    } else if(j > 3){ //skip col 0 and 3 which are ID and Email
                        Cell currentCell = targRow.createCell(j-1);
                        currentCell.setCellValue(srcCell.getStringCellValue());
                        newSheet.autoSizeColumn(j);//FIX: Not resizing columns...
                    }
                }else{
                    //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
                }
            }
        }
        
        //STEP 3: Copy new week into Generated Report sheet
        
        Sheet oldGen = connReport.getSheet("Current Report");
        connReport.setSheetName(2, "Old Report");
        Sheet newGen = connReport.createSheet("Current Report");
        newGen = connReport.cloneSheet(2);
        connReport.setSheetOrder("Current Report", 2);
        
        
        //One Way: but List<> cannot be initialized
//        ArrayList<Row> rowCopier = new ArrayList<Row>() {};
//        newSheet.copyRows(rowCopier, 0, new CellCopyPolicy());
        
//        for(Cell cell : row){
//            
//        }
//        Cell cell = row.getCell(2);
//
//        // Create the cell if it doesn't exist
//        if (cell == null)
//            cell = row.createCell(2);
//
//        // Update the cell's value
//        cell.setCellType(CellType.STRING);
//        cell.setCellValue("Updated Value");

        // Write the output to the file
        FileOutputStream fileOut = new FileOutputStream("ConReport041618.xlsx");
        connReport.write(fileOut);
        
        //attempt to fix autoSizeColumn. There was a stackexchange saying it needed to be after .write
        for(int i = 0; i < 11; i++){
            newSheet.autoSizeColumn(i);
        }
        fileOut.close();

        // Closing the workbook
        connReport.close();
    }
    
    
    
}

//Solution for autosize
/*
XSSFFont font = workbook.createFont();
font.setFontName("Arial");
XSSFCellStyle style = workbook.createCellStyle();
style.setFont(font);

//autosize doesn't help with calibri font
*/