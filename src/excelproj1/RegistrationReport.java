/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author WoodmDav
 */
public class RegistrationReport {
    
    public static void executeAutomation(File weeklyRegReport, File ytdRegCSV, File updatedRegReport) {
        try {
            System.out.println("Setting up Registration Report");
            String dateToday = myTools.getDate();
            ZipSecureFile.setMinInflateRatio(0);
            
            myTools.csvToXLSX(weeklyRegReport, "src//Resources//DatabaseRecords//", "Weekly_Reg_Report" + myTools.getWeek() + ".xlsx");
            myTools.csvToXLSX(ytdRegCSV, "src//Resources//DatabaseRecords//", "YTD_Reg_Report" + myTools.getWeek() + ".xlsx");
            XSSFWorkbook regReport = new XSSFWorkbook(updatedRegReport);
            XSSFWorkbook weeklyReg = new XSSFWorkbook(new File("src//Resources//DatabaseRecords//Weekly_Reg_Report" + myTools.getWeek() + ".xlsx"));
            XSSFWorkbook YTDReg = new XSSFWorkbook(new File("src//Resources//DatabaseRecords//YTD_Reg_Report" + myTools.getWeek() + ".xlsx"));

            //idk if i need the next line
            regReport.setForceFormulaRecalculation(true);//recalculate all formuals upon opening

            //Establish old week and new week to prepare to copy headers
            regReport.createSheet(dateToday);//getDate() gives you Jul 07 or something like it
            regReport.setSheetOrder(dateToday, 3);//move new sheet in front of other weeks

            XSSFSheet newYTDsheet = regReport.getSheetAt(3);//grab the new sheet
            XSSFSheet lastWeekSheet = regReport.getSheetAt(4);//grab previous week
            XSSFSheet updatedWeeklySheet = regReport.getSheetAt(1);//grab the cumulative weekly sheet
            XSSFSheet cumulativeSheet = regReport.getSheetAt(2);//grab the Cumulative Report sheet
            XSSFSheet ytdSrcSheet = YTDReg.getSheetAt(0);//Grab the only sheet in YTD Reg Report
            XSSFSheet weeklySrcSheet = weeklyReg.getSheetAt(0);//grab only sheet in weekly




            System.out.println("Step One: New YTD Sheet");
            ytdSheet(regReport, lastWeekSheet, newYTDsheet, ytdSrcSheet, dateToday);

            System.out.println("Step Two: Update Weekly Sheet");
            updateWeekly(updatedWeeklySheet, weeklySrcSheet, regReport);


            System.out.println("Step Three: Cumulative Report");
            updateCumulative(cumulativeSheet, dateToday);



            //save
            System.out.print("Registration Report Complete! Saving...");
//            FileOutputStream fileOut = new FileOutputStream("src//Resources//Reports//YTD Updated Registration Report " + myTools.getWeek() + ".xlsx");
            FileOutputStream tempFileOut = new FileOutputStream("src//Resources//NewFiles//YTD Updated Registration Report " + myTools.getWeek() + ".xlsx");
//            regReport.write(fileOut);
            regReport.write(tempFileOut);
            regReport.close();
            System.out.println(" Saved!");
            
        } catch (IOException ex) {
            Logger.getLogger(RegistrationReport.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(RegistrationReport.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private static void ytdSheet(XSSFWorkbook regReport, Sheet lastWeekSheet, Sheet newYTDsheet, Sheet ytdSrcSheet, String dateToday){
        //copy header
        myTools.copyCells(regReport, lastWeekSheet, newYTDsheet, 0, 0, 2, 0, 0, 0);
        //copy data
        myTools.copyCells(regReport, ytdSrcSheet, newYTDsheet, 0, 1, 1, ytdSrcSheet.getPhysicalNumberOfRows(), 1, 1);
        //build dates
        for(int row = newYTDsheet.getPhysicalNumberOfRows(); row > 0; row--){//start at the end of sheet and move upward//optim:
            try{
                Cell targCell = newYTDsheet.getRow(row).getCell(1);
                //System.out.println("targcellenum: " + targCell.getCellTypeEnum());
                //System.out.println((targCell.getCellTypeEnum().toString().equals("NUMERIC")));
                if(targCell.getCellTypeEnum().toString().equals("STRING")||targCell.getCellTypeEnum().toString().equals("NUMERIC")){//if you find cell with data, then start copying dates
                    newYTDsheet.getRow(row).createCell(0).setCellValue(dateToday);
                }
            }catch(NullPointerException e){
                //System.out.println("shouldn't be null, check new week date");
                //null cell
            }
        }
        int firstNullYTD = myTools.findFirstNullRow(newYTDsheet, 1);
        for(int row = 1; row < firstNullYTD; row++){
            Cell dateCol = newYTDsheet.getRow(row).createCell(0);
            dateCol.setCellValue(dateToday);//set date of first column

            Cell regNum = newYTDsheet.getRow(row).getCell(2);
            int regAsInt = Integer.parseInt(regNum.getStringCellValue());
            regNum.setCellType(CellType.NUMERIC); //grab reg numbers make sure they're written as a number
            regNum.setCellValue(regAsInt);//for some reason setting cell type was changing the value. here's my solution.
        }
        //total formula
        Row totalRow = newYTDsheet.createRow(firstNullYTD+1); //skip a row
        totalRow.createCell(1).setCellValue("TOTAL: ");
        totalRow.createCell(2).setCellFormula("SUM(C2:C" + firstNullYTD + ")");

    }
    
    private static void updateWeekly(XSSFSheet updatedWeeklySheet, XSSFSheet weeklySrcSheet, XSSFWorkbook workbook){
        //look for the last row with data in it, and get the index
//        int rowStartIndex = (myTools.findFirstNullRow(updatedWeeklySheet, 7)); //subtract 1 to override one day of the previous week
           

        int colStartIndex = 2; //Latest Date Column
        int colEndIndex = (myTools.findFirstNullRow(updatedWeeklySheet, 3));
//        for(int i = 0; i < 8; i++){
//            //updatedWeeklySheet.shiftColumns(colStartIndex, colEndIndex, 1);            BROKEN
//            //gets higher each time the data shifts right
//            
//            colEndIndex++;
//        }
      

        myTools.shiftColumns(workbook, updatedWeeklySheet, 2, 2, 4, 8);
        
        Row dateRow = updatedWeeklySheet.getRow(2);//Date Row
        Row regRow = updatedWeeklySheet.getRow(3);//Registrations Row
        for(int row = 9; row > 0; row--){//grab from weekly.xlsx and leave out header
            try{
                //d/System.out.println("row: " + row);
                //d/System.out.println("rowStartIndex: " + rowStartIndex);
                
//                Row targRow = updatedWeeklySheet.createRow(rowStartIndex++);
                XSSFRow srcRow = weeklySrcSheet.getRow(row);
                int targCol = row + 1;
                XSSFCell srcCell = srcRow.getCell(0);
                System.out.println("Cell: " + srcCell.getStringCellValue());
                if(srcCell.getStringCellValue().length() > 0){
                    dateRow.createCell(targCol).setCellValue(srcCell.getStringCellValue());

                    srcCell = srcRow.getCell(1);
//                    double regAsInt = (srcCell.getNumericCellValue());
//                    srcCell.setCellType(CellType.NUMERIC); //grab reg numbers make sure they're written as a number
//                    srcCell.setCellValue(regAsInt);//for some reason setting cell type was changing the value. here's my solution.
                    
                    
                    DataFormatter formatter = new DataFormatter();
                    int val = Integer.parseInt(formatter.formatCellValue(srcCell));
                    System.out.println("val: " + val);
                    regRow.createCell(targCol).setCellValue(val);     
                    
                    //TEMP
//                    XSSFRow targRow = updatedWeeklySheet.getRow(5);
//                    targRow.createCell(targCol).setCellFormula("FLOOR(AVERAGE(" + myTools.indexToLetter(targCol) + " 4:" + myTools.indexToLetter(targCol + 7) + " 4),1)");
                }
                targCol = 3;
                //AVG
                XSSFRow targRow = updatedWeeklySheet.getRow(5);
                targRow.createCell(targCol).setCellFormula("FLOOR(AVERAGE(" + myTools.indexToLetter(targCol) + "4:" + myTools.indexToLetter(targCol + 7) + "4),1)");
                //MAX
                targRow = updatedWeeklySheet.getRow(6);
                targRow.createCell(targCol).setCellFormula("MAX(" + myTools.indexToLetter(targCol) + "4:" + myTools.indexToLetter(targCol + 7) + "4)");
                //SUM
                targRow = updatedWeeklySheet.getRow(7);
                targRow.createCell(targCol).setCellFormula("SUM(" + myTools.indexToLetter(targCol) + "4:" + myTools.indexToLetter(targCol + 7) + "4)");
            }catch(NullPointerException e){
//                e.printStackTrace();
                System.out.println("NULL POINTER: unexpected null at rowVar - " + row);
            }
        }
    }
    
    private static void updateCumulative(XSSFSheet cumulativeSheet, String dateToday){
        
//        int firstNullColumn = myTools.findFirstNullColumn(cumulativeSheet, 1);
        int firstNullColumn = 9; //column J
        int cumulativeSheetRowCells = cumulativeSheet.getRow(6).getPhysicalNumberOfCells();
        System.out.println("Cumulative row cell count: " + cumulativeSheetRowCells);
        cumulativeSheet.shiftColumns(firstNullColumn, cumulativeSheetRowCells, 1);
        String columnIndex = myTools.indexToLetter(firstNullColumn);
        cumulativeSheet.getRow(1).createCell(firstNullColumn).setCellValue(dateToday);
        int row;
        for(row = 2; row < 14; row++){
            //build model formulas, need to change the number to reflect the actual row number, not index
            cumulativeSheet.getRow(row).createCell(firstNullColumn).setCellFormula("SUMIF('" + dateToday + "'!$B$2:$B$100,($I" + (row + 1) + "&\"*\"),'" + dateToday + "'!$C$2:$C$100)");
        }
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellValue("0");
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM(" + columnIndex + "3:" + columnIndex + "15)");
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM('" + dateToday + "'!C:C)/2");

        row++;//skip row
        
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellValue(dateToday);
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM(" + columnIndex + "6:" + columnIndex + "14)");
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM(" + columnIndex + "4:" + columnIndex + "5)");
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM(" + columnIndex + "3:" + columnIndex + "3)");
        cumulativeSheet.getRow(row++).createCell(firstNullColumn).setCellFormula("SUM(" + columnIndex + "15)");
        //percent values at B2:B5
        cumulativeSheet.getRow(1).createCell(1).setCellFormula("J20");
        cumulativeSheet.getRow(2).createCell(1).setCellFormula("J21");
        cumulativeSheet.getRow(3).createCell(1).setCellFormula("J22");
    }
}
