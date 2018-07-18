/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelproj1;

import java.text.SimpleDateFormat;
import java.util.Date;
import static javafx.scene.input.KeyCode.M;

/**
 *
 * @author WoodmDav
 */
public class Testing {
    public static void main(String[] args){
        
//        System.out.println(26/27);
//        System.out.println(28/27);
//        System.out.println(60/27);

        int a = 65;
        int b = 25+65;
        int c = 60%27 + 65;
        
//        char charA = (char)a;
//        char charB = (char)b;
//        char charC = (char)c;
//        
//        System.out.println(charA);
//        System.out.println(charB);
//        System.out.println(charC);
        
        for(int i = 0; i < (26*27); i++){
            System.out.println("Result: " + indexToLetter(i));
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
        StringBuffer buff = new StringBuffer( );
        String month = monthFormat.format(today);
        String day = dayFormat.format(today);
        
        System.out.println(month + " " + day);
    }
}
