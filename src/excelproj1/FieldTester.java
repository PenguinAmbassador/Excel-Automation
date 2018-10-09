package excelproj1;

import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author WoodmDav
 */
public class FieldTester {
    private String name;
    private String model;
    private String serial;
    private String mac;
    private String origFW;
    private String currentFW;
    private String type;
    
    public FieldTester(Row row){
        name = row.getCell(3).getStringCellValue() + " " + row.getCell(2).getStringCellValue();
        serial = row.getCell(6).getStringCellValue();
        currentFW = row.getCell(7).getStringCellValue();
        model = row.getCell(8).getStringCellValue();
        mac = row.getCell(9).getStringCellValue();
        origFW = row.getCell(10).getStringCellValue();
        
        //declare type based on first four letters of model
        System.out.println("Type maybe: " + model.substring(0, 3));
        if(model.substring(0,4).equals("ENGH") || model.substring(0,3).equals("FGVH")){
            type = "2-in-1";
        }else if(model.substring(0,4).equals("FFRE") || model.substring(0,3).equals("FGRC")){
            type = "Radical RAC";
        }else if(model.substring(0,4).equals("FGPC")){
            type = "Stromboli RAC";
        }else if(model.substring(0,4).equals("FGAC")){
            type = "Dehumidifier";
        }else{
            System.out.println("Model " + model + " is not yet supported by class FieldTester");
        }
    }
    
    public void printFieldTester(){
        System.out.println("Name: " + name);
        System.out.println("Model: " + model);
        System.out.println("Serial: " + serial);
        System.out.println("MAC: " + mac);
        System.out.println("OrigFW: " + origFW);
        System.out.println("CurrentFW: " + currentFW);
        System.out.println("Type: " + type);
        System.out.println("");
    }

    public String getName() {
        return name;
    }

    public String getModel() {
        return model;
    }

    public String getSerial() {
        return serial;
    }

    public String getMac() {
        return mac;
    }

    public String getOrigFW() {
        return origFW;
    }

    public String getCurrentFW() {
        return currentFW;
    }

    public String getType() {
        return type;
    }
    
    
}
