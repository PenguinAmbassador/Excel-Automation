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
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class findNewFieldTesters {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook connReport = WorkbookFactory.create(new File("src\\Updated Connectivity Report 05-28-18.xlsx"));
        Sheet targSheet = connReport.getSheet("FT Participants");
        Sheet srcSheet = connReport.getSheet("Current Report");
        myTools.searchColumn(srcSheet, 7, 2, 90, targSheet, 4, 15, 211);
        
    }
}