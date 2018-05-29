/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

import java.io.File;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
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
                    cloneCellStringValue(oldCell,newCell);
                }
            }else{//if shiftCount is negative push left
                //TODO create loop to push cells left
                System.out.println("You are trying to push cells left, right? Sorry, that code doesn't exist at the moment!");
            }
        }
    }
    //will this work between two different excel workbooks? Please test.
    //srcSheet and targSheet will be the same if you are working within one sheet
    void copyRows(Sheet srcSheet,Sheet targSheet, int rowStartIndex, int rowEndIndex, int colStartIndex, int colEndIndex){
        for(int i = rowStartIndex; i < rowEndIndex; i++){ 
            Row srcRow = srcSheet.getRow(i); 
            Row targRow;
            if(targSheet.getRow(i)==null){
                targRow = targSheet.createRow(i); 
            }else{
                targRow = targSheet.getRow(i);
            }
            for(int j = colStartIndex; j < colEndIndex; j++){
                Cell srcCell = srcRow.getCell(j); //track source cell in Weekly Report
                if(srcCell!=null && targRow.getCell(j)!=null){ //none are null
                    Cell targCell = targRow.getCell(j);
                    cloneCellStringValue(srcCell, targCell);
                }else if(srcCell!=null && targRow.getCell(j)==null){//target is null
                    Cell targCell = targRow.createCell(j);
                    cloneCellStringValue(srcCell, targCell);
                }else{//both are null
                    //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
                }
            }
        }
    }
    
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
    void cloneCellStringValue(Cell oldCell, Cell newCell){
        newCell.setCellValue(oldCell.getStringCellValue());
        //there is a better version of this on stack exchange
    }
}
