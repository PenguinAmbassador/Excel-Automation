package excelproj1;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import static java.util.Objects.isNull;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *Excel Utility Class
 * @author WoodmDav
 */
public class myTools {
    
    /**
     * Finds the last column with any data within a row and returns the column number after 
     * @param targetSheet
     * @param targetRow
     * @return 
     */
    public static int findFirstNullColumn(Sheet targetSheet, int targetRow){
        for(int col = (targetSheet.getRow(targetRow).getPhysicalNumberOfCells() + 50); col > 0; col--){//start at the end of sheet and move upward
            try{
                Cell targCell = targetSheet.getRow(targetRow).getCell(col);
                //System.out.println("targcellenum: " + targCell.getCellTypeEnum());
                //System.out.println((targCell.getCellTypeEnum().toString().equals("NUMERIC")));
                if(targCell.getCellTypeEnum().toString().equals("STRING")||targCell.getCellTypeEnum().toString().equals("NUMERIC")){//if you find cell with data, then return the row after it
                    return ++col;
                }
            }catch(NullPointerException e){
                //System.out.println("shouldn't be null, check new week date");
                //null cell
            }
        }
        System.out.println("Unable find data. MyTools.findFirstNull() error.");
        return -1;
    }
    
    /**
     * Finds the last row with any data within a column and returns the row number after
     * @param targetSheet
     * @param targetColumn
     * @return 
     */
    public static int findFirstNullRow(Sheet targetSheet, int targetColumn){
        for(int rowIterator = targetSheet.getPhysicalNumberOfRows(); rowIterator > 0; rowIterator--){//start at the end of sheet and move upward
            try{
                Cell targCell = targetSheet.getRow(rowIterator).getCell(targetColumn);
                //System.out.println("targcellenum: " + targCell.getCellTypeEnum());
                //System.out.println((targCell.getCellTypeEnum().toString().equals("NUMERIC")));
                if(targCell.getCellTypeEnum().toString().equals("STRING")||targCell.getCellTypeEnum().toString().equals("NUMERIC")){//if you find cell with data, then return the row after it
                    return ++rowIterator;
                }
            }catch(NullPointerException e){
                //null cell
            }
        }
        System.out.println("Unable find data. MyTools.findFirstNull() error.");
        return -1;
    }
    
    /**
     * POI extension by Sankumarsingh stck. Saves csv file to new location based on parameters
     * @param targetFile
     * @param newXlsxPath
     * @param newXslxName 
     */
    public static void csvToXLSX(File targetFile, String newXlsxPath, String newXslxName) {
        try {
            //String csvFileAddress = path; //csv file address
            String xlsxFileAddress = newXlsxPath + newXslxName; //xlsx file address
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("sheet1");
            String currentLine=null;
            int rowNum=0;
            BufferedReader br = new BufferedReader(new FileReader(targetFile));
            while ((currentLine = br.readLine()) != null) {
                String[] str = currentLine.split(",");
                XSSFRow currentRow=sheet.createRow(rowNum++);
                for(int i=0;i<str.length;i++){
                    currentRow.createCell(i).setCellValue(str[i]);
                }
            }

            FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
            System.out.println("FAIL: MyTools - csvToXLSX");
        }
    }
    
    /**
     * Return a code representative of the index of a cell
     * @param index 
     * @return 0 becomes A 27 becomes AA 28 becomes AB and so on
     */
    public static String indexToLetter(int index){
        String let = "";
        char tempChar;
        if((index/26)>0){                                       //26 letters in the alphabet. 
            tempChar = (char)(index/26 + 64);                   //64 is 1 less so that we start with AA instead of BA
            let = let.concat(Character.toString(tempChar));     //first letter
            tempChar = (char)(index%26 + 65);                   //65 is an offset to get to the first letter in asci
            let = let.concat(Character.toString(tempChar));     //second letter
        }else if (index > 600){                                 
            System.out.println("ERROR: indexToLetter does not support index higher than 600");
        }else{
            tempChar = (char)(index + 65); 
            let = Character.toString(tempChar);
        }
        return let;
    }
    
    /**
     * For every cell in this range set the cell value to "" empty string
     * @param sheet
     * @param colStart
     * @param rowStart
     * @param colEnd
     * @param rowEnd 
     */
    public static void deleteCells(Sheet sheet, int colStart, int rowStart, int colEnd, int rowEnd){
        for(int row = rowStart; row < rowEnd; row++){
            for(int cell = colStart; cell < colEnd;cell++){
                try{
                    Row tempRow = sheet.getRow(row);  
                    Cell tempCell = tempRow.getCell(cell);
                    tempCell.setCellValue("");
                }catch(NullPointerException e){
                    //e.printStackTrace();
                    //d/System.out.println("Failed to delete cell: " + cell + "," + row);
                }
            }
        }
    }
        
    //You should plug in sheet.getPhysicalNumberOfRows as rowEndIndex if you want to iterate all rows.
    //ultimately a useless method... there is already a shift method in poi
    public static void shiftColumns(Workbook workbook, Sheet sheet, int cellIndex, int rowStartIndex, int rowEndIndex, int shiftCount) {
        for(int i = rowStartIndex; i < rowEndIndex; i++){//This should iterate through rows specified by parameters
            Row currentRow = sheet.getRow(i);
            if(shiftCount > 0){//if shiftCount is positive push right
                for (int j = currentRow.getPhysicalNumberOfCells()-1;j>=cellIndex;j--){//iterate starting on the right and push cells right
                    Cell oldCell = currentRow.getCell(j);
                    //First line creates cell, second line copies old cell into new cell
                    Cell newCell = currentRow.createCell(j + shiftCount);//let it be noted that I removed second argument oldCell.getCellTypeEnum()
                    
                    if(oldCell != null){
                        cloneCellValue(oldCell, newCell, workbook);
                    }
                }
            }else{//if shiftCount is negative push left
                //TODO create loop to push cells left
                System.out.println("You are trying to push cells left, right? Sorry, that code doesn't exist at the moment!");
            }
        }
    }
    
    /**
     * Copy cells from one sheet to another in the ranges specified
     * @param workbook The workbook you are copying within
     * @param srcSheet
     * @param targSheet
     * @param fromColStartIndex
     * @param fromRowStartIndex
     * @param fromColEndIndex
     * @param fromRowEndIndex
     * @param toColStartIndex
     * @param toRowStartIndex 
     */
    static void copyCells(XSSFWorkbook workbook, Sheet srcSheet,Sheet targSheet, int fromColStartIndex, int fromRowStartIndex, int fromColEndIndex, int fromRowEndIndex, int toColStartIndex, int toRowStartIndex){
        int rowOffset = (fromRowStartIndex - toRowStartIndex); //offset difference between target and source
        int cellOffset = (fromColStartIndex - toColStartIndex); //offset difference between target and source
        
        //for 
        for(int row = toRowStartIndex; row <= (toRowStartIndex + (fromRowEndIndex - fromRowStartIndex)); row++){//loop from start to end of the range
            Row srcRow = srcSheet.getRow(row + rowOffset); //copy from here
            Row targRow;//copy to here
            if(targSheet.getRow(row)==null){
                targRow = targSheet.createRow(row);
            }else{
                targRow = targSheet.getRow(row);
            }
            //d/System.out.println("row: " + row + " rowOffSet: " + (row+rowOffset));
            for(int col = toColStartIndex; col <= (toColStartIndex + (fromColEndIndex - fromColStartIndex)); col++){//loop from start to end of the range
                try{
                    Cell srcCell = srcRow.getCell(col + cellOffset); //track source cell in Weekly Report
                    if(srcCell!=null && targRow.getCell(col)!=null){ //none are null
                        //d/System.out.println("row/col: " + row + "/"  + col + "colOffset: " + (col+cellOffset));
                        Cell targCell = targRow.getCell(col);//offest for where you aim to put the copied cells
                        cloneCellValue(srcCell, targCell, workbook);
                    }else if(srcCell!=null && targRow.getCell(col)==null){//target is null
                        Cell targCell = targRow.createCell(col);
                        cloneCellValue(srcCell, targCell, workbook);
                    }else{//src is empty or both are null
                        //blank cells. ultimately needs to be fixed by a fillBlanks() method that pulls names from emails
                    }
                }catch(NullPointerException e){
                    //e.printStackTrace();
                    //d/System.out.println("null at row/col: " + row + "/" + col);
                }
                
            }
        }
    }
    
    
    static void searchColumn(Sheet targSheet, int targColumnIndex, int targRowStartIndex, int targRowEndIndex, Sheet srcSheet, int srcColumnIndex, int srcRowStartIndex, int srcRowEndIndex){
        for(int i = targRowStartIndex; i < targRowEndIndex; i++){ 
            Row targRow = targSheet.getRow(i);
            //colIndex let's us look at a specific column through specified range
            Cell targCell = targRow.getCell(targColumnIndex);
            boolean newFieldTester = true;
//            if(targCell!=null){
//                System.out.println("ROW " + (i+1) + ": " + targCell.getStringCellValue());
//            }
            for(int j = srcRowStartIndex; (j < srcRowEndIndex) && newFieldTester; j++){//if they are an old field tester exit the loop and go to the next row.
                Row srcRow = srcSheet.getRow(j);
                Cell srcCell = srcRow.getCell(srcColumnIndex);
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
                System.out.println("Found: " + targCell.getStringCellValue());
            }
        }
    }
    
    static ArrayList<FieldTester> findFieldTesters(Sheet currentReport, int targColumnIndex, int targRowStartIndex, int targRowEndIndex, Sheet ftSheet, int srcColumnIndex, int srcRowStartIndex, int srcRowEndIndex){
        ArrayList<FieldTester> fieldTesters=  new ArrayList<>();
        for(int i = targRowStartIndex; i < targRowEndIndex; i++){
            if(isNull(currentReport.getRow(i))){
                //d/System.out.println("Empty Row: " + i);
            }else{
                Row targRow = currentReport.getRow(i);
                //colIndex let's us look at a specific column through specified range
                Cell targCell = targRow.getCell(targColumnIndex);
                boolean newFieldTester = true;
                for(int j = srcRowStartIndex; (j < srcRowEndIndex) && newFieldTester; j++){//if they are an old field tester exit the loop and go to the next row.
                    try{
                        //d/System.out.println("Row: " + j);
                        //d/System.out.println("Col: " + srcColumnIndex);
                        Row srcRow = ftSheet.getRow(j);
                        Cell srcCell = srcRow.getCell(srcColumnIndex);
                        //d/System.out.println(srcCell.getStringCellValue());
                        boolean cellsExist = srcCell != null && targCell !=null;
                        boolean cellsMatch = cellsExist && (srcCell.getStringCellValue().trim().equalsIgnoreCase(targCell.getStringCellValue().trim()));
                        newFieldTester = cellsExist && !cellsMatch; //if cellsExist and cells don't match then field tester is new
                    }catch(NullPointerException e ){
                        //d/System.out.println("null cell");
                        //d/System.out.println("row: " + i);
                        //d/System.out.println("col: " + srcColumnIndex);
                    }
                }
                if(newFieldTester){
                    fieldTesters.add(new FieldTester(targRow));
                }
            }
            
        }
        for(int i = 0; i < fieldTesters.size(); i++){
            fieldTesters.get(i).printFieldTester();
        }
        return fieldTesters;
    } 
    
    /**
     * Copies values from one cell to the next. Currently only supports string/numeric/formula
     * @param oldCell
     * @param newCell
     * @param srcWkBk 
     */
    static void cloneCellValue(Cell oldCell, Cell newCell, Workbook srcWkBk){

        CellStyle newCellStyle = srcWkBk.createCellStyle();
        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
        newCell.setCellStyle(newCellStyle);
        
        //System.out.println(oldCell.getCellTypeEnum().toString());
        try{
            //if it's a string, number, or formula copy data accordingly
            if(oldCell.getCellTypeEnum().toString().equals("STRING")){
                newCell.setCellValue(oldCell.getStringCellValue());
            } else if(oldCell.getCellTypeEnum().toString().equals("NUMERIC")){
                newCell.setCellValue(oldCell.getNumericCellValue());
            }else if(oldCell.getCellTypeEnum().toString().equals("FORMULA")){
                newCell.setCellType(CellType.FORMULA);
                newCell.setCellFormula(oldCell.getCellFormula());
                //System.out.println(oldCell.getCellFormula()+ " other: " + newCell.getCellFormula());
            }
        }catch(IllegalStateException e){
            //for some reason oldCell had a String Formula? Setting cell type got rid of this error.
            newCell.setCellType(CellType.FORMULA);
            newCell.setCellFormula(oldCell.getStringCellValue());
            e.printStackTrace();
            System.out.println("Illegal State Exception");
            System.out.println(oldCell.getStringCellValue());
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
    
    //TODO get ride of over usage by saving variable once on run.
    /**
    * Returns the date in format of "Jul 06"
    * @return 
    */
    public static String getDate(){
        Date today = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM-dd-yy");
        String date = dateFormat.format(today);
        return(date); 
    }
    
    /**
    * Returns the date of last week in the form of "07-16-18"
    * @return 
    */
    public static String getLastWeek(){
        long DAY_IN_MS = 1000 * 60 * 60 * 24;
        Date lastWeek = new Date(System.currentTimeMillis() - (7 * DAY_IN_MS));
        SimpleDateFormat monthFormat = new SimpleDateFormat("MM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        SimpleDateFormat yearFormat = new SimpleDateFormat("YY");
        String month = monthFormat.format(lastWeek);
        String day = dayFormat.format(lastWeek);
        String year = yearFormat.format(lastWeek);
        
        return(month + "-" + day + "-" + year); 
    }
    /**
    * Returns the date of last week in the form of "07-16-18"
    * @return 
    */
    public static String getWeek(){
        Date lastWeek = new Date();
        SimpleDateFormat monthFormat = new SimpleDateFormat("MM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        SimpleDateFormat yearFormat = new SimpleDateFormat("YY");
        String month = monthFormat.format(lastWeek);
        String day = dayFormat.format(lastWeek);
        String year = yearFormat.format(lastWeek);
        
        return(month + "-" + day + "-" + year); 
    }
}
