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
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * TODO: VERIFY TOTALS, ADD FORMULA OPTION TO GUI?
 * @author WoodmDav
 */
public class ConnectivityReport {
    
    
    public String NEW_NIUX_FW;
    public String NEW_GEN2_FW;
    
    ConnectivityReport(File weeklyFT, File connReport, Gui gui, String niuxFW, String gen2FW){        
        this.NEW_NIUX_FW = niuxFW;
        this.NEW_GEN2_FW = gen2FW;
        executeAutomation(weeklyFT, connReport);
    }
    
    public void executeAutomation(File weeklyFT, File connReport){
        try {
            System.out.println("Setting up Connectivity Report");
            ZipSecureFile.setMinInflateRatio(0);//prevented a zip error
            XSSFWorkbook connectivityWorkbook = new XSSFWorkbook(connReport);//Obtain a Connectivity Report from xlsx; this report should be a week older than the weekly report
            myTools.csvToXLSX(weeklyFT, "src//Resources//DatabaseRecords//", "Weekly_Report" + myTools.getWeek() + ".xlsx");//csv files must be converted to xlsx
            XSSFWorkbook weeklyWorkbook = new XSSFWorkbook(new File("src//Resources//DatabaseRecords//Weekly_Report" + myTools.getWeek() + ".xlsx"));//Obtain Weekly Report. Used to build Connectivity Report.
            connectivityWorkbook.setForceFormulaRecalculation(true);//recalculate all formulas upon opening
            
            //Establish old week and new week to prepare to copy headers
            connectivityWorkbook.createSheet(myTools.getDate());//getDate() gives you Jul 07 or something like it
            connectivityWorkbook.setSheetOrder(myTools.getDate(), 4);//move new sheet in front of other weeks
            XSSFSheet newWeekSheet = connectivityWorkbook.getSheetAt(4);//grab the new sheet
            XSSFSheet lastWeekSheet = connectivityWorkbook.getSheetAt(5);//grab previous week
            XSSFSheet weeklySheet = weeklyWorkbook.getSheetAt(0);//Grab the only sheet in Weekly Report
            XSSFSheet currentWeekSheet = connectivityWorkbook.getSheet("Current Report");
            
            
            System.out.println("Step One: New Week Sheet");
            buildNewWeek(newWeekSheet, lastWeekSheet, weeklySheet);
            
            System.out.println("Step Two: Current Report");
            updateCurrentReport(connectivityWorkbook, currentWeekSheet, newWeekSheet);

            System.out.println("Step Three: Generated Report");
            buildGeneratedReport(connectivityWorkbook);

            //Step 4: FT stuff TODO-learned it make more sense to be able to take a sheet pass it into a  new object and give it methods like sheet.copyCells
            System.out.println("Step Four: FT Sheet");
            newFieldTrialColumn(connectivityWorkbook, currentWeekSheet);



            System.out.print("Report Complete! Saving...");
            // Write the output to the file            
            
//            FileOutputStream fileOut = new FileOutputStream("src//Resources//Reports//Connectivity Report " + myTools.getWeek() + ".xlsx");
            FileOutputStream tempFileOut = new FileOutputStream("src//Resources//NewFiles//Connectivity Report " + myTools.getWeek() + ".xlsx");
            
//            connectivityWorkbook.write(fileOut);
            connectivityWorkbook.write(tempFileOut);
            tempFileOut.close();
            connectivityWorkbook.close(); // Closing the workbook
            System.out.println(" Saved!\n");
            
        } catch (IOException ex) {
            Logger.getLogger(ConnectivityReport.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("IOException reading Connectivity Report");
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ConnectivityReport.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("InvalidFormatException reading Connectivity Report");
        }
    }
    
    private static void updateCurrentReport(XSSFWorkbook connReport, Sheet currentReport, Sheet newSheet){
//        System.out.println("Comment this the next line unless testing");
//        Sheet newSheet = connReport.getSheet("Jul 9");
        myTools.deleteCells(currentReport, 0, 1, 12, currentReport.getPhysicalNumberOfRows()); //delete old report
        myTools.copyCells(connReport, newSheet, currentReport, 0, 1, 23, newSheet.getPhysicalNumberOfRows(), 1,1);//copy from the new week
        
        //start at the end of sheet and move upward
        for(int row = currentReport.getPhysicalNumberOfRows(); row > 0; row--){
            try{
                Cell targCell = currentReport.getRow(row).getCell(1);
                //System.out.println("targcellenum: " + targCell.getCellTypeEnum());
                //System.out.println((targCell.getCellTypeEnum().toString().equals("NUMERIC")));
                if(targCell.getCellTypeEnum().toString().equals("NUMERIC")){//if you find cell with data, then start copying dates
                    currentReport.getRow(row).createCell(0).setCellValue(myTools.getDate());
                }
            }catch(Exception NullPointerException){
                //null cell
            }
        }
    }
    
    private static void addFieldTrialParticipants(Sheet currentReport, Sheet ftSheet){
    //myTools.searchColumn(currentReport, 7, 2, 90, ftSheet, 4, 15, 211);
        ArrayList<FieldTester> fieldTesters = myTools.findFieldTesters(currentReport, 6, 1, 90, ftSheet, 3, 24, 215);
        for(FieldTester dude : fieldTesters){
                    
            for(int row = 24; row < ftSheet.getPhysicalNumberOfRows(); row++){//checking for type of field tester
                try{
                    Row targRow = ftSheet.getRow(row);
                    //d/System.out.println("row: " + row);
                    
                    //d/System.out.println("Title: " + targRow.getCell(0).getStringCellValue());
                    //d/System.out.println("FT Type: " + dude.getType());
                    if(targRow.getCell(0).getStringCellValue().equals(dude.getType())){
                        ftSheet.shiftRows(++row, ftSheet.getPhysicalNumberOfRows(), 1);//insert row after the first row matching the type
                        targRow = ftSheet.createRow(row);//create row on newly created row
                        
                        int cell = 1;
                        Cell tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getName());
                        tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getModel());
                        tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getSerial());
                        tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getMac());
                        tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getOrigFW());
                        tempCell = targRow.createCell(cell++);
                        tempCell.setCellValue(dude.getCurrentFW());
                        
                        row = ftSheet.getPhysicalNumberOfRows();
                    }
                }catch(NullPointerException e){
                    System.out.println("null: " + row);
                }
                
            }
        }
    }
    
    private static void newFieldTrialColumn(Workbook connReport, Sheet currentReport){
        Sheet ftSheet = connReport.getSheet("FT Participants");
        myTools.shiftColumns(connReport, ftSheet, 7, 23, ftSheet.getPhysicalNumberOfRows()-1, 1);//TODO this line corrupts data is there are blanks while shifting
        
        //Time to make new column
        ftSheet.getRow(23).getCell(7).setCellValue(myTools.getDate()); //setdate
        String formula = "IF(ISNA(VLOOKUP($D25,'" + myTools.getDate() + "'!$F:$F,1,0)),\"No\",\"Yes\")";
        for(int row = 24; row < ftSheet.getPhysicalNumberOfRows(); row++){
            Row targRow = ftSheet.getRow(row);
            Cell targCell = targRow.getCell(7);
            try{
                targCell.setCellFormula("IF(ISNA(VLOOKUP($D" + (row + 1) + ",'" + myTools.getDate() + "'!$F:$F,1,0)),\"No\",\"Yes\")");
            }catch(NullPointerException e){
                //nullcells
            }
        }
        
        addFieldTrialParticipants(currentReport, ftSheet);
    }
    
    private static void buildNewWeek(Sheet newSheet, Sheet lastWeekSheet, Sheet weeklySheet){
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
    }
    
    private void buildGeneratedReport(XSSFWorkbook connReport){
        XSSFSheet genRep = connReport.getSheet("Generated Report");//store generated report
        
//        //4.1: GET COL INDEX
//        int newColIndex = 0;
//        Row targRow = genRep.getRow(12);// row index 8 is where the data tables start
//        //Going to count down from the end of the row moving to the left until I hit a formula.
//        //Once I don't have any errors, then that means I am on the last column and can add a new column to the right
//        for(int col = targRow.getPhysicalNumberOfCells()+50; col >= 0; col--){
//            try{
//                //System.out.println("check1: " + col);
//                Cell targCell = targRow.getCell(col);
//                String checkForm = targCell.getCellFormula();
//                //d/System.out.println(checkForm);
//            if(checkForm.length() > 0){
//                newColIndex = col + 1;
//                col = -1;//stop loop
//            }
//            }catch(Exception NullPointerException){
//                //null cell
//            }
//        }
        
        int newColIndex = 9;
        
        int rowStart = 11;
        int rowEnd = 23;
//        
//        myTools.shiftColumns(connReport, genRep, newColIndex, rowStart, rowEnd, 1);   //newColIndex, genRep.getRow(11), 1);
//        myTools.shiftColumns(connReport, genRep, newColIndex, 29, 30, 1); 
//        myTools.shiftColumns(connReport, genRep, newColIndex, 33, 35, 1); 
//        myTools.shiftColumns(connReport, genRep, newColIndex, 39, 64, 1); 

        genRep.shiftColumns(newColIndex, genRep.getRow(15).getPhysicalNumberOfCells(), 1);

//        myTools.shiftColumns( genRep, newColIndex, rowStart, rowEnd, 1);   //newColIndex, genRep.getRow(11), 1);
//        myTools.shiftColumns(genRep.getRow(i), newColIndex, 1
//        )newColIndex, 29, 30, 1); 
//        myTools.shiftColumns(connReport, genRep, newColIndex, 33, 35, 1); 
//        myTools.shiftColumns(connReport, genRep, newColIndex, 39, 64, 1); 
        
        //4.2 COPY column over
        ConnFormula form = new ConnFormula(newColIndex, NEW_NIUX_FW, NEW_GEN2_FW);
        int rowIndex = 11;
        Row targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex).setCellValue(myTools.getDate());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getFRIDGE_NEW());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getFRIDGE_OLD());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getRAC_X_NEW());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getRAC_X_OLD());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getRAC_GEN2_NEW());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getRAC_GEN2_OLD());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getSTROMBO_NEW());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getSTROMBO_OLD());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getDEHUM_NEW());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getDEHUM_OLD());
        
        rowIndex++;//skip a row
        
        targRow = genRep.getRow(rowIndex);//row index 23
        rowIndex++;
        targRow.createCell(newColIndex).setCellValue(myTools.getDate());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getFRIDGE_TOTAL());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getRAC_TOTAL());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex);
        targRow.getCell(newColIndex).setCellFormula(form.getSTROMBO_TOTAL());
        targRow = genRep.getRow(rowIndex);
        rowIndex++;
        targRow.createCell(newColIndex).setCellFormula(form.getDEHUM_TOTAL());
        
        rowIndex++;//skip a row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getTARGET());
        
        rowIndex++;//skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getTOTAL());

        rowIndex++; //skip row  
        rowIndex++;
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getOLDFW_TOTAL());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getNEWFW_TOTAL());

        rowIndex++;//skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellValue(myTools.getDate());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getFRIDGE_FW1_PW1MA076());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getFRIDGE_FW2_PW1MA079());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW1_PW1RS326());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW2_v4310());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW3_v4420());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW4_v453b());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW5_v4551());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW6_v4642());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW7_v4852());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW8_v49653());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getRAC_FW9_v4107());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW1_PW3RS017_161005a());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW2_v4310());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW3_v4420());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW4_v453b());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW5_v4551());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW6_v4642());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW7_v4852());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW8_v49653());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getSTROMBO_FW9_v4107());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW1_v4310());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW2_v4420());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW3_v453b());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW4_v4551());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW5_v4642());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW6_v4852());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW7_v49653());
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getDEHUM_FW8_v4107());
        
        rowIndex++; //skip row
        
        genRep.getRow(rowIndex++).createCell(newColIndex).setCellFormula(form.getFW_TOTAL());
        
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