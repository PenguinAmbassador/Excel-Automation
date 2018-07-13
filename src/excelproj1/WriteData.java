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
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import tutorial.Employee;

/**
 * TODO: VERIFY TOTALS, ADD FORMULA OPTION TO GUI?,AUTOMATICALLY ADD FIELD TESTERS
 * @author WoodmDav
 */
public class WriteData {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        System.out.println("Running Main WriteData");
        // Obtain a Connectivity Report from xlsx
        //This report should be a week older than the weekly report
        Workbook connReportOld = WorkbookFactory.create(new File("src\\ConnectivityReport07-02-18.xlsx"));
        Workbook connReport = connReportOld;//delete this... didn't work
        //Obtain Weekly Report. Used to build Connectivity Report.
        Workbook weeklyReport = WorkbookFactory.create(new File("src\\Weekly_Report07-09-18.xlsx"));
        
//        try{
//            if(connReport.getSheetAt(4).getSheetName().equals("Jul 13")){
//                connReport.removeSheetAt(4); //Delete this try/catch later, just making it so i can test faster
//            }
//        }catch(Exception e){
//        }
        
        //Establish old week and new week to prepare to copy headers
//        String newSheetName = myTools.getDate();//getDate() gives you Jul 07 or something like it
//        connReport.createSheet(newSheetName);
//        connReport.setSheetOrder(myTools.getDate(), 4);//move new sheet in front of other weeks
//        Sheet newSheet = connReport.getSheetAt(4);//grab the new sheet
//        Sheet lastWeekSheet = connReport.getSheetAt(5);//grab previous week
//        Sheet weeklySheet = weeklyReport.getSheetAt(0);//Grab the only sheet in Weekly Report
//        
//        // STEP 1: Copy headers from an old week 
//        Row rowSrc = lastWeekSheet.getRow(0);
//        Row rowDest = newSheet.createRow(0);
//        int colNum = 0;//helps iterate through the columns
//        for(Cell cell : rowSrc){
//            Cell currentCell = rowDest.createCell(colNum);
//            currentCell.setCellValue(cell.getStringCellValue());
//            currentCell.setCellStyle(cell.getCellStyle());
//            newSheet.autoSizeColumn(colNum);//FIX: Not resizing columns...
//            colNum++;
//        }
//        
//        //STEP 2: Copy data from Weekly Report
//        for(int i = 1; i < weeklySheet.getPhysicalNumberOfRows(); i++){ //copy every row but row 1
//            //System.out.println("i: " + i);
//            Row srcRow = weeklySheet.getRow(i); //Track source row in Weekly Report
//            Row targRow = newSheet.createRow(i); //Track target row in newSheet in the Connectivity report
//            for(int j = 0; j < 12; j++){ //copy every cell
//                //System.out.println("j: " + j);
//                //System.out.println(srcRow.getCell(j));
//                Cell srcCell = srcRow.getCell(j); //track source cell in Weekly Report
//                if(srcCell!=null){
//                    if(j == 0){
//                        Cell currentCell = targRow.createCell(j);
//                        currentCell.setCellValue(i); //store row as #
//                    }
//                    if(j > 0 && j < 3){ //skip col 0 and 3 which are ID and Email
//                        Cell currentCell = targRow.createCell(j);
//                        currentCell.setCellValue(srcCell.getStringCellValue());
//                        newSheet.autoSizeColumn(j);//FIX: Not resizing columns...
//                    } else if(j > 3){ //skip col 0 and 3 which are ID and Email
//                        Cell currentCell = targRow.createCell(j-1);
//                        currentCell.setCellValue(srcCell.getStringCellValue());
//                        newSheet.autoSizeColumn(j);//FIX: Not resizing columns...
//                    }
//                }else{
//                    //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
//                }
//            }
//        }
//        
//
//        System.out.println("Step Three");
//        
//        //STEP 3: Copy new week into Generated Report sheet
//        //Sheet newSheet = connReport.getSheet(getDate());
//        Sheet CurrentReport = connReport.getSheet("Current Report");
//        
//        //I want to copy cells A2 to K2 all the way through to bottom
//        //COPY FROM:  A2 = row0col1 -> K2 = 11, 1 COPY TO: B2 row1col1
//        myTools.copyCells(newSheet, CurrentReport, 0, 1, 11, newSheet.getPhysicalNumberOfRows(), 1,1);
        //TODO: Generated report doesn't have dates copying over.

        //Step 4: formula junk
        Sheet genRep = connReport.getSheet("Generated Report");//store generated report
        System.out.println("Step 4");
        
        //4.1: GET COL INDEX
        int newColIndex = 0;
        Row targRow = genRep.getRow(12);// row index 8 is where the data tables start
        //Going to count down from the end of the row moving to the left until I hit a formula.
        //Once I don't have any errors, then that means I am on the last column and can add a new column to the right
        for(int col = targRow.getPhysicalNumberOfCells()+50; col >= 0; col--){
            try{
                System.out.println("check1: " + col);
                Cell targCell = targRow.getCell(col);
                String checkForm = targCell.getCellFormula();
                System.out.println(checkForm);
            if(checkForm.length() > 0){
                newColIndex = col + 1;
                col = -1;//stop loop
            }
            }catch(Exception NullPointerException){
                //null cell
            }
        }
        
        //4.2 COPY column over
        System.out.println("check3: " + newColIndex);
        int rowIndex = 11;
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex).setCellValue(myTools.getDate());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.FRIDGE_NEW);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.FRIDGE_OLD);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.RAC_X_NEW);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.RAC_X_OLD);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.RAC_GEN2_NEW);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.RAC_GEN2_OLD);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.STROMBO_NEW);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.STROMBO_OLD);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.DEHUM_NEW);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.DEHUM_OLD);
        
        rowIndex++;//skip a row
        
        targRow = genRep.getRow(rowIndex);//row index 23
        rowIndex++;
        targRow.createCell(newColIndex).setCellValue(myTools.getDate());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.FRIDGE_TOTAL);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.RAC_TOTAL);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(Constants.STROMBO_TOTAL);
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex).setCellFormula(Constants.DEHUM_TOTAL);
        
        rowIndex++;//skip a row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.TARGET);
        
        rowIndex++;//skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.TOTAL);
        
        rowIndex++; //skip row  
        rowIndex++;
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.OLDFW_TOTAL);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.NEWFW_TOTAL);
        
        rowIndex++;//skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.FRIDGE_FW1_PW1MA076);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.FRIDGE_FW2_PW1MA079);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW1_PW1RS326);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW2_v4310);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW3_v4420);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW4_v453b);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW5_v4551);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.RAC_FW6_v4642);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW1_PW3RS017_161005a);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW2_v4310);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW3_v4420);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW4_v453b);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW5_v4551);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.STROMBO_FW6_v4642);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.DEHUM_FW1_v4310);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.DEHUM_FW2_v4420);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.DEHUM_FW3_v453b);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.DEHUM_FW4_v4551);
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.DEHUM_FW5_v4642);
        
        rowIndex++; //skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(Constants.FW_TOTAL);
        
        
        
        
        
        
        
        
        
        // Write the output to the file
        FileOutputStream fileOut = new FileOutputStream("ConReport042318 5.xlsx");
        connReport.write(fileOut);
        
//        //attempt to fix autoSizeColumn. There was a stackexchange saying it needed to be after .write
//        for(int i = 0; i < 11; i++){
//            newSheet.autoSizeColumn(i);
//        }
//        fileOut.close();

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