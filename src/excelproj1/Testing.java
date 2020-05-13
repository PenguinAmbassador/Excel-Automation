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
import java.util.Scanner;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.logging.Level;
import java.util.logging.Logger;
//weather
import com.google.gson.*;
import com.google.gson.reflect.*;


/**
 *
 * @author WoodmDav
 */
public class Testing {
    public static void main(String[] args){
        
        //RegistrationReport.executeAutomation(new File("src//Weekly_Reg_Report.csv"), new File("src//YTD_Reg_Report.csv"), new File("src//YTD Updated Registration Report 07-16-18.xlsx"));\
        
        String API_KEY = "7d0b1a990aae61051e20071727a86053";
        String LOCATION = "7d0b1a990aae61051e20071727a86053";
        
//        String urlString = ""
        
        
        
        
        
        
        
        
//        boolean isMetric = true;
//        String owmApiKey = "XXXXXXXXXXXX"; /* YOUR OWM API KEY HERE */
//        String weatherCity = "Brisbane,AU";
//        byte forecastDays = 3;
//        OpenWeatherMap.Units units = (isMetric)
//            ? OpenWeatherMap.Units.METRIC
//            : OpenWeatherMap.Units.IMPERIAL;
//        OpenWeatherMap owm = new OpenWeatherMap(units, owmApiKey);
//        try {
//          DailyForecast forecast = owm.dailyForecastByCityName(weatherCity, forecastDays);
//          System.out.println("Weather for: " + forecast.getCityInstance().getCityName());
//          int numForecasts = forecast.getForecastCount();
//          for (int i = 0; i < numForecasts; i++) {
//            DailyForecast.Forecast dayForecast = forecast.getForecastInstance(i);
//            DailyForecast.Forecast.Temperature temperature = dayForecast.getTemperatureInstance();
//            System.out.println("\t" + dayForecast.getDateTime());
//            System.out.println("\tTemperature: " + temperature.getMinimumTemperature() +
//                " to " + temperature.getMaximumTemperature() + "\n");
//          }
//        }
//        catch (IOException | JSONException e) {
//          e.printStackTrace();
//        }
  
}
 

 
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
