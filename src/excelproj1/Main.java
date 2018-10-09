package excelproj1;

import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author WoodmDav
 */
public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException{
        
        System.out.println("MAIN RUN");
        jFrame gui = new jFrame();
        gui.setTitle("Excel Automation Engine");
        gui.setVisible(true);
        
        //ConnectivityReport.executeAutomation();
        //RegistrationReport.executeAutomation();
    }
}
