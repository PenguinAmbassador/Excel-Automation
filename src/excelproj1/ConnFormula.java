package excelproj1;

/**
 *
 * @author WoodmDav
 */
class ConnFormula {

    private String NEW_NIUX_FW;
    private String NEW_GEN2_FW;
    private String FRIDGE_NEW;
    private String FRIDGE_OLD;
    private String RAC_X_NEW;
    private String RAC_X_OLD;
    private String RAC_GEN2_NEW;
    private String RAC_GEN2_OLD;
    private String STROMBO_NEW;
    private String DEHUM_NEW;
    private String DEHUM_OLD;
    private String STROMBO_OLD;
    private String FRIDGE_TOTAL;
    private String STROMBO_TOTAL;
    private String DEHUM_TOTAL;
    private String RAC_TOTAL;
    private String TARGET;
    private String TOTAL;
    private String OLDFW_TOTAL;
    private String NEWFW_TOTAL;
    private String FRIDGE_FW1_PW1MA076;
    private String FRIDGE_FW2_PW1MA079;
    private String RAC_FW1_PW1RS326;
    private String RAC_FW3_v4420;
    private String RAC_FW4_v453b;
    private String RAC_FW2_v4310;
    private String RAC_FW5_v4551;
    private String RAC_FW6_v4642;
    private String RAC_FW7_v4852;
    private String RAC_FW8_v49653;
    private String RAC_FW9_v4107;
    private String STROMBO_FW1_PW3RS017_161005a;
    private String STROMBO_FW2_v4310;
    private String STROMBO_FW3_v4420;
    private String STROMBO_FW4_v453b;
    private String STROMBO_FW5_v4551;
    private String STROMBO_FW6_v4642;
    private String STROMBO_FW7_v4852;
    private String STROMBO_FW8_v49653;
    private String STROMBO_FW9_v4107;
    private String DEHUM_FW2_v4420;
    private String DEHUM_FW4_v4551;
    private String DEHUM_FW3_v453b;
    private String DEHUM_FW1_v4310;
    private String DEHUM_FW5_v4642;
    private String DEHUM_FW6_v4852;
    private String DEHUM_FW7_v49653;
    private String DEHUM_FW8_v4107;
    private String FW_TOTAL;
    
    ConnFormula(int index, String niuxFW, String gen2FW){
        
        String colLetter = myTools.indexToLetter(index);//represents turning column index 0 into A, column 26 into AA, and so forth
        String today = myTools.getDate();
                
        NEW_NIUX_FW = niuxFW;
        NEW_GEN2_FW = gen2FW; 
        //I replaced date with a function and I replaced a reference to new fw with a constants 
        FRIDGE_NEW = "COUNTIFS('" + today + "'!$G:$G,\"PW1MA079*\")+COUNTIFS('" + today + "'!$G:$G,\"PW1MA076*\")";
        FRIDGE_OLD = "(COUNTIF('" + today + "'!$G:$G,\"PW1MA*\")+COUNTIF('" + today + "'!$G:$G,\"PW3MA*\"))-" + colLetter + "13";//can potentially replace 13 with param aswell
        RAC_X_NEW = "COUNTIFS('" + today + "'!$G:$G, \"" + NEW_NIUX_FW + "\" ,'" + today + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + today + "'!$G:$G, \"" + NEW_NIUX_FW + "\" ,'" + today + "'!$H:$H,\"FFRE*\")";
        RAC_X_OLD = "(COUNTIFS('" + today + "'!$G:$G,\"v*\",'" + today + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + today + "'!$G:$G,\"PW3RS*\",'" + today + "'!$H:$H,\"FFRE*\"))-" + colLetter + "15";
        RAC_GEN2_NEW = "COUNTIFS('" + today + "'!$G:$G,\"" + NEW_GEN2_FW + "\",'" + today + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + today + "'!$G:$G,\"" + NEW_GEN2_FW + "\",'" + today + "'!$H:$H,\"FFRE*\")";
        RAC_GEN2_OLD = "(COUNTIFS('" + today + "'!$G:$G,\"PW1RS*\",'" + today + "'!$H:$H,\"FFRE*\")+COUNTIFS('" + today + "'!$G:$G,\"PW1RS*\",'" + today + "'!$H:$H,\"FGRC*\"))-" + colLetter + "17";
        STROMBO_NEW = "COUNTIFS('" + today + "'!$G:$G, \"" + NEW_NIUX_FW + "\" ,'" + today + "'!$H:$H,\"FGPC*\")";
        STROMBO_OLD = "COUNTIFS('" + today + "'!$H:$H,\"FGPC*\")-" + colLetter + "19";
        DEHUM_NEW = "COUNTIFS('" + today + "'!$G:$G, \"" + NEW_NIUX_FW + "\" ,'" + today + "'!$H:$H,\"FGAC*\")";
        DEHUM_OLD = "COUNTIFS('" + today + "'!$H:$H,\"FGAC*\")-" + colLetter + "21";

        FRIDGE_TOTAL = "SUM(" + colLetter + "$13:" + colLetter + "$14)";
        RAC_TOTAL = "SUM(" + colLetter + "$15:" + colLetter + "$18)";
        STROMBO_TOTAL = "SUM(" + colLetter + "$19:" + colLetter + "$20)";
        DEHUM_TOTAL = "" + colLetter + "21+" + colLetter + "22";
       
        TARGET = "$A$72";
        TOTAL = "SUM(" + colLetter + "$25:" + colLetter + "$28)";
       
        OLDFW_TOTAL = "SUM(" + colLetter + "$14," + colLetter + "$16," + colLetter + "$18," + colLetter + "$20," + colLetter + "$22)";
        NEWFW_TOTAL = "SUM(" + colLetter + "$13," + colLetter + "$15," + colLetter + "$17," + colLetter + "$19," + colLetter + "$21)";

        //TODO this can obviously can be made into a method... can the rest?
        int row = 41;
        FRIDGE_FW1_PW1MA076 = "COUNTIFS('" + today + "'!$G:$G, $I" + row++ + "&\"*\")";
        FRIDGE_FW2_PW1MA079 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\")";
        RAC_FW1_PW1RS326 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\")";
        RAC_FW2_v4310 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW3_v4420 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW4_v453b = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW5_v4551 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW6_v4642 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW7_v4852 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW8_v49653 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        RAC_FW9_v4107 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGRC\"&\"*\")";
        STROMBO_FW1_PW3RS017_161005a = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW2_v4310 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW3_v4420 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW4_v453b = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW5_v4551 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW6_v4642 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW7_v4852 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW8_v49653 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        STROMBO_FW9_v4107 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGPC\"&\"*\")";
        DEHUM_FW1_v4310 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW2_v4420 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW3_v453b = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW4_v4551 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW5_v4642 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW6_v4852 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW7_v49653 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        DEHUM_FW8_v4107 = "COUNTIFS('" + today + "'!$G:$G,$I" + row++ + "&\"*\", '" + today + "'!$H:$H, \"FGAC\"&\"*\")";
        FW_TOTAL ="SUM(" + colLetter + "41:" + colLetter + (--row) + ")";

    }

    public String getRAC_FW9_v4107() {
        return RAC_FW9_v4107;
    }

    public String getSTROMBO_FW9_v4107() {
        return STROMBO_FW9_v4107;
    }

    public String getDEHUM_FW8_v4107() {
        return DEHUM_FW8_v4107;
    }    
    
    public String getRAC_FW8_v49653() {
        return RAC_FW8_v49653;
    }

    public String getSTROMBO_FW8_v49653() {
        return STROMBO_FW8_v49653;
    }

    public String getDEHUM_FW7_v49653() {
        return DEHUM_FW7_v49653;
    }

    public void setNEW_NIUX_FW(String NEW_NIUX_FW) {
        this.NEW_NIUX_FW = NEW_NIUX_FW;
    }

    public void setNEW_GEN2_FW(String NEW_GEN2_FW) {
        this.NEW_GEN2_FW = NEW_GEN2_FW;
    }
    
    public String getNewestFW() {
        return NEW_NIUX_FW;
    }

    public String getNewishFW() {
        return NEW_GEN2_FW;
    }

    public String getFRIDGE_NEW() {
        return FRIDGE_NEW;
    }

    public String getFRIDGE_OLD() {
        return FRIDGE_OLD;
    }

    public String getRAC_X_NEW() {
        return RAC_X_NEW;
    }

    public String getRAC_X_OLD() {
        return RAC_X_OLD;
    }

    public String getRAC_GEN2_NEW() {
        return RAC_GEN2_NEW;
    }

    public String getRAC_GEN2_OLD() {
        return RAC_GEN2_OLD;
    }

    public String getSTROMBO_NEW() {
        return STROMBO_NEW;
    }

    public String getDEHUM_NEW() {
        return DEHUM_NEW;
    }

    public String getDEHUM_OLD() {
        return DEHUM_OLD;
    }

    public String getSTROMBO_OLD() {
        return STROMBO_OLD;
    }

    public String getFRIDGE_TOTAL() {
        return FRIDGE_TOTAL;
    }

    public String getSTROMBO_TOTAL() {
        return STROMBO_TOTAL;
    }

    public String getDEHUM_TOTAL() {
        return DEHUM_TOTAL;
    }

    public String getRAC_TOTAL() {
        return RAC_TOTAL;
    }

    public String getTARGET() {
        return TARGET;
    }

    public String getTOTAL() {
        return TOTAL;
    }

    public String getOLDFW_TOTAL() {
        return OLDFW_TOTAL;
    }

    public String getNEWFW_TOTAL() {
        return NEWFW_TOTAL;
    }

    public String getFRIDGE_FW1_PW1MA076() {
        return FRIDGE_FW1_PW1MA076;
    }

    public String getFRIDGE_FW2_PW1MA079() {
        return FRIDGE_FW2_PW1MA079;
    }

    public String getRAC_FW1_PW1RS326() {
        return RAC_FW1_PW1RS326;
    }

    public String getRAC_FW3_v4420() {
        return RAC_FW3_v4420;
    }

    public String getRAC_FW4_v453b() {
        return RAC_FW4_v453b;
    }

    public String getRAC_FW2_v4310() {
        return RAC_FW2_v4310;
    }

    public String getRAC_FW5_v4551() {
        return RAC_FW5_v4551;
    }

    public String getRAC_FW6_v4642() {
        return RAC_FW6_v4642;
    }

    public String getSTROMBO_FW1_PW3RS017_161005a() {
        return STROMBO_FW1_PW3RS017_161005a;
    }

    public String getSTROMBO_FW2_v4310() {
        return STROMBO_FW2_v4310;
    }

    public String getSTROMBO_FW3_v4420() {
        return STROMBO_FW3_v4420;
    }

    public String getSTROMBO_FW4_v453b() {
        return STROMBO_FW4_v453b;
    }

    public String getSTROMBO_FW5_v4551() {
        return STROMBO_FW5_v4551;
    }

    public String getSTROMBO_FW6_v4642() {
        return STROMBO_FW6_v4642;
    }

    public String getDEHUM_FW2_v4420() {
        return DEHUM_FW2_v4420;
    }

    public String getDEHUM_FW4_v4551() {
        return DEHUM_FW4_v4551;
    }

    public String getDEHUM_FW3_v453b() {
        return DEHUM_FW3_v453b;
    }

    public String getDEHUM_FW1_v4310() {
        return DEHUM_FW1_v4310;
    }

    public String getDEHUM_FW5_v4642() {
        return DEHUM_FW5_v4642;
    }

    public String getNEW_NIUX_FW() {
        return NEW_NIUX_FW;
    }

    public String getNEW_GEN2_FW() {
        return NEW_GEN2_FW;
    }

    public String getRAC_FW7_v4852() {
        return RAC_FW7_v4852;
    }

    public String getSTROMBO_FW7_v4852() {
        return STROMBO_FW7_v4852;
    }

    public String getDEHUM_FW6_v4852() {
        return DEHUM_FW6_v4852;
    }

    public String getFW_TOTAL() {
        return FW_TOTAL;
    }
    
}
