/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author WoodmDav
 */
public class myTools {
    //You should plug in sheet.getPhysicalNumberOfRows as rowEndIndex if you want to iterate all rows.
    void shiftColumns(Sheet sheet, int rowStartIndex, int rowEndIndex, int cellIndex, int shiftCount) {
        for(int i = rowStartIndex; i < rowEndIndex; i++){//This should iterate through rows specified by parameters
            Row currentRow = sheet.getRow(i);
            if(shiftCount > 0){//if shiftCount is positive push right
                for (int j = currentRow.getPhysicalNumberOfCells()-1;j>=cellIndex;j--){//iterate starting on the right and push cells right
                    Cell oldCell = currentRow.getCell(j);
                    //First line creates cell, second line copies old cell into new cell
                    Cell newCell = currentRow.createCell(j + shiftCount);//let it be noted that I removed second argument oldCell.getCellTypeEnum()
                    cloneCellValue(oldCell,newCell);
                }
            }else{//if shiftCount is negative push left
                //TODO create loop to push cells left
                System.out.println("You are trying to push cells left, right? Sorry, that code doesn't exist at the moment!");
            }
        }
    }
    //will this work between two different excel workbooks? Please test.
    //srcSheet and targSheet will be the same if you are working within one sheet
    static void copyCells(Sheet srcSheet,Sheet targSheet, int fromColStartIndex, int fromRowStartIndex, int fromColEndIndex, int fromRowEndIndex, int toColStartIndex, int toRowStartIndex){
        //Every time we get a cell from old sheet, we need to account for the offset between old and new sheets based on starting places. 
        int rowOffset = (fromRowStartIndex - toRowStartIndex);
        int cellOffset = (fromColStartIndex - toColStartIndex);
        
        for(int row = toRowStartIndex; row < (toRowStartIndex + (fromRowEndIndex - fromRowStartIndex)); row++){  
            Row srcRow = srcSheet.getRow(row + rowOffset); 
            Row targRow;
            if(targSheet.getRow(row)==null){
                System.out.println("");
                targRow = targSheet.createRow(row);
            }else{
                targRow = targSheet.getRow(row);
            }
            System.out.println("row: " + row + " rowOffSet: " + (row+rowOffset));
            for(int col = toColStartIndex; col < (toColStartIndex + (fromColEndIndex - fromColStartIndex)); col++){
                try{
                    Cell srcCell = srcRow.getCell(col + cellOffset); //track source cell in Weekly Report
                    if(srcCell!=null && targRow.getCell(col)!=null){ //none are null
                        System.out.println("row/col: " + row + "/"  + col + "colOffset: " + (col+cellOffset));
                        Cell targCell = targRow.getCell(col);//offest for where you aim to put the copied cells
                        cloneCellValue(srcCell, targCell);
                    }else if(srcCell!=null && targRow.getCell(col)==null){//target is null
                        Cell targCell = targRow.createCell(col);
                        cloneCellValue(srcCell, targCell);
                    }else{//both are null
                        //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
                    }
                }catch(Exception e){
                    System.out.println("Failed at i/j " + row + "/" + col);
                    e.printStackTrace();
                }
                
            }
        }
    }
    
    //Not sure if this actually works! I just removed srcSheet param
    static void copyCells(Sheet targSheet, int cellStartIndex, int rowEndIndex, int colStartIndex, int colEndIndex){
        for(int i = cellStartIndex; i < rowEndIndex; i++){ 
            Row srcRow = targSheet.getRow(i); 
            Row targRow;
            if(targSheet.getRow(i)==null){
                targRow = targSheet.createRow(i); 
            }else{
                targRow = targSheet.getRow(i);
            }
            for(int j = colStartIndex; j < colEndIndex; j++){
                
                try{
                    Cell srcCell = srcRow.getCell(j);
                    if(srcCell!=null && targRow.getCell(j)!=null){ //none are null
                    Cell targCell = targRow.getCell(j);
                    cloneCellValue(srcCell, targCell);
                    }else if(srcCell!=null && targRow.getCell(j)==null){//target is null
                        Cell targCell = targRow.createCell(j);
                        cloneCellValue(srcCell, targCell);
                    }else{//both are null
                        //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
                    }
                } catch(Exception e){
                    System.out.println("Failed at i/j " + i + "/" + j);
                    e.printStackTrace();
                }
                
            }
        }
    }
    
    //TODO I really wanna replace this logic with functionality based on REAL index
    //you can enter the exact row/col numbers and method will adjust for the index.
    static void searchColumn(Sheet targSheet, int targColumnIndex, int targRowStartIndex, int targRowEndIndex, Sheet srcSheet, int srcColumnIndex, int srcRowStartIndex, int srcRowEndIndex){
        for(int i = targRowStartIndex-1; i < targRowEndIndex; i++){ 
            Row targRow = targSheet.getRow(i);
            //colIndex let's us look at a specific column through specified range
            Cell targCell = targRow.getCell(targColumnIndex-1);
            boolean newFieldTester = true;
//            if(targCell!=null){
//                System.out.println("ROW " + (i+1) + ": " + targCell.getStringCellValue());
//            }
            for(int j = srcRowStartIndex-1; (j < srcRowEndIndex) && newFieldTester; j++){//if they are an old field tester exit the loop and go to the next row.
                Row srcRow = srcSheet.getRow(j);
                Cell srcCell = srcRow.getCell(srcColumnIndex-1);
//                if(srcCell != null){
//                    System.out.println("ROW " + (j+1) + ": " + srcCell.getStringCellValue());
//                }
                boolean cellsExist = srcCell != null && targCell !=null;
                boolean cellsMatch = cellsExist && (srcCell.getStringCellValue().trim().equalsIgnoreCase(targCell.getStringCellValue().trim()));
//                System.out.println(cellsExist);
//                System.out.println(cellsMatch);
                newFieldTester = cellsExist && !cellsMatch; //if cellsExist and cells don't match then field tester is new
//                System.out.println(newFieldTester);
            }
//            System.out.println("ENDJLOOP");
            if(newFieldTester){
//                System.out.println("SUCCESS!");
                System.out.println(targCell.getStringCellValue());
            }
        }
    }
    //NEXT TODO: create copyCell, copyRow, copyHeader
    /**
     * Copies values from one cell to the next. Currently only supports string/numeric
     * @param oldCell
     * @param newCell 
     */
    static void cloneCellValue(Cell oldCell, Cell newCell){
        System.out.println(oldCell.getCellTypeEnum().toString());
        try{
            if(oldCell.getCellTypeEnum().toString().equals("STRING")){
                newCell.setCellValue(oldCell.getStringCellValue());
            } else if(oldCell.getCellTypeEnum().toString().equals("NUMERIC")){
                newCell.setCellValue(oldCell.getNumericCellValue());
            }else if(oldCell.getCellTypeEnum().toString().equals("FORMULA")){
                newCell.setCellValue(oldCell.getCellFormula());
            }
        }catch(Exception e){
            e.printStackTrace();
            System.out.println("ERR CAUGHT: MyTools cloneCellValue() - data type may not yet be supported");
        }
        
        //there is a better version of this on stack exchange
    }
    static void fixNames(Sheet weeklyReport){
        String firstName = "";
        String lastName = "";
        for(int i = 0; i < weeklyReport.getPhysicalNumberOfRows(); i++){
            //check if the names are empty (col Index 1 & 2)
            Row currentRow = weeklyReport.getRow(i);
            for(int j = 1; j <= 2; j++){
                Cell currentCell = currentRow.getCell(j);
                if(currentCell==null && (j==1)){
                    //get first name from email (col index 3)
                    int k = 0;
                    while(currentCell.getStringCellValue().charAt(k)!= '.'){//grab all letters before .
                        String newLetter = new StringBuilder().append(currentCell.getStringCellValue().charAt(k++)).toString();//POSSIBLE PROBLEM your k++; grab char, turn to string with strBuilder
                        firstName = firstName.concat(newLetter);
                    }
                } else if(currentCell==null && (j==2)){
                    //get first name from email (col index 3)
                    int k = 0;
                    while(currentCell.getStringCellValue().charAt(k)!= '@'){//grab all letters before @
                        String newLetter = new StringBuilder().append(currentCell.getStringCellValue().charAt(k++)).toString();//grab char, turn to string with strBuilder
                        lastName = lastName.concat(newLetter);
                    }
                }
            }
            //TODO: save firstlastname to col indices 1&2
        }
    }
    
    /**
    * Returns the date in format of "Jul 06"
    * @return 
    */
    public static String getDate(){
        Date today = new Date();
        SimpleDateFormat monthFormat = new SimpleDateFormat("MMM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        String month = monthFormat.format(today);
        String day = dayFormat.format(today);

        return(month + " " + day); 
    }
}
