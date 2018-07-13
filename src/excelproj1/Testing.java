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
        Date today = new Date();
        SimpleDateFormat monthFormat = new SimpleDateFormat("MMM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        StringBuffer buff = new StringBuffer( );
        String month = monthFormat.format(today);
        String day = dayFormat.format(today);
        
        System.out.println(month + " " + day);
    }
}
