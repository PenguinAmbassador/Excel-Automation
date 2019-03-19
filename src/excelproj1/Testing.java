/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author WoodmDav
 */
public class Testing {
    public static void main(String[] args){
        
        RegistrationReport.executeAutomation(new File("src//Weekly_Reg_Report.csv"), new File("src//YTD_Reg_Report.csv"), new File("src//YTD Updated Registration Report 07-16-18.xlsx"));
    }
    
    public static String indexToLetter(int index){
        String let = "";
        char tempChar;
        if((index/26)>0){
            tempChar = (char)(index/26 + 64);//1 less so that we start with AA instead of BA
            let = let.concat(Character.toString(tempChar));
            tempChar = (char)(index%26 + 65);
            let = let.concat(Character.toString(tempChar));
        }else{
            tempChar = (char)(index + 65); 
            let = Character.toString(tempChar);
        }
        return let;
    }
    
    private static void dateTest(){
        
        Date today = new Date();
        SimpleDateFormat monthFormat = new SimpleDateFormat("MMM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        String month = monthFormat.format(today);
        String day = dayFormat.format(today);
        
        System.out.println(month + " " + day);
    }
    
    private static void findFieldTesters() throws IOException, InvalidFormatException{
        Workbook connReport = WorkbookFactory.create(new File("src\\Updated Connectivity Report 05-28-18.xlsx"));
        Sheet targSheet = connReport.getSheet("FT Participants");
        Sheet srcSheet = connReport.getSheet("Current Report");
        myTools.searchColumn(srcSheet, 6, 1, 90, targSheet, 3, 14, 211); //i broke this when I tried to get rid of index logic
    }
}
