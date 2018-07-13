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
public class Constants {
    public static String newestFW = "v4.6-4-2";
    public static String newishFW = "PW1RS326"; //newish meaning it's the newest FW for NIU gen2
    
    //date
    //I replaced date with a function and I replaced a reference to new fw with a constants 
    public static final String FRIDGE_NEW = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1MA079*\")+COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1MA076*\")";
    public static final String FRIDGE_OLD = "(COUNTIF('" + myTools.getDate() + "'!$G:$G,\"PW1MA*\")+COUNTIF('" + myTools.getDate() + "'!$G:$G,\"PW3MA*\")) - (" + FRIDGE_NEW + ")";
    public static final String RAC_X_NEW = "COUNTIFS('" + myTools.getDate() + "'!$G:$G, \"" + newestFW + "\" ,'" + myTools.getDate() + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + myTools.getDate() + "'!$G:$G, \"" + newestFW + "\" ,'" + myTools.getDate() + "'!$H:$H,\"FFRE*\")";
    public static final String RAC_X_OLD = "(COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"v*\",'" + myTools.getDate() + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW3RS*\",'" + myTools.getDate() + "'!$H:$H,\"FFRE*\"))- (" + RAC_X_NEW + ")";
    public static final String RAC_GEN2_NEW = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1RS326\",'" + myTools.getDate() + "'!$H:$H,\"FGRC*\")+COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1RS326\",'" + myTools.getDate() + "'!$H:$H,\"FFRE*\")";
    public static final String RAC_GEN2_OLD = "(COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1RS*\",'" + myTools.getDate() + "'!$H:$H,\"FFRE*\")+COUNTIFS('" + myTools.getDate() + "'!$G:$G,\"PW1RS*\",'" + myTools.getDate() + "'!$H:$H,\"FGRC*\"))- (" + RAC_GEN2_NEW + ")";
    public static final String STROMBO_NEW = "COUNTIFS('" + myTools.getDate() + "'!$G:$G, \"" + newestFW + "\" ,'" + myTools.getDate() + "'!$H:$H,\"FGPC*\")";
    public static final String STROMBO_OLD = "COUNTIFS('" + myTools.getDate() + "'!$H:$H,\"FGPC*\")- (" + STROMBO_NEW + ")";
    public static final String DEHUM_NEW = "COUNTIFS('" + myTools.getDate() + "'!$G:$G, \"" + newestFW + "\" ,'" + myTools.getDate() + "'!$H:$H,\"FGAC*\")";
    public static final String DEHUM_OLD = "COUNTIFS('" + myTools.getDate() + "'!$H:$H,\"FGAC*\")- (" + DEHUM_NEW + ")";
    
    //date
    public static final String FRIDGE_TOTAL = FRIDGE_NEW +" + "+ FRIDGE_OLD;
    public static final String RAC_TOTAL = RAC_X_NEW +" + "+ RAC_X_OLD +" + "+ RAC_GEN2_NEW +" + "+ RAC_GEN2_OLD;
    public static final String STROMBO_TOTAL = STROMBO_OLD+" + "+STROMBO_NEW;
    public static final String DEHUM_TOTAL = DEHUM_NEW+" + "+DEHUM_OLD;
    //date
    public static final String TARGET = "$A$68";
    public static final String TOTAL = FRIDGE_TOTAL+" + "+RAC_TOTAL+" + "+STROMBO_TOTAL+" + "+DEHUM_TOTAL;
    //date
    public static final String OLDFW_TOTAL = FRIDGE_OLD+" + "+RAC_X_OLD+" + "+RAC_GEN2_OLD+" + "+STROMBO_OLD+" + "+DEHUM_OLD;
    public static final String NEWFW_TOTAL = FRIDGE_NEW+" + "+RAC_X_NEW+" + "+RAC_GEN2_NEW+" + "+STROMBO_NEW+" + "+DEHUM_NEW;
    
    //date
    public static final String FRIDGE_FW1_PW1MA076 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G, $I41&\"*\")";
    public static final String FRIDGE_FW2_PW1MA079 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I42&\"*\")";
    public static final String RAC_FW1_PW1RS326 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I43&\"*\")";
    public static final String RAC_FW2_v4310 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I44&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGRC\"&\"*\")";
    public static final String RAC_FW3_v4420 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I45&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGRC\"&\"*\")";
    public static final String RAC_FW4_v453b = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I46&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGRC\"&\"*\")";
    public static final String RAC_FW5_v4551 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I47&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGRC\"&\"*\")";
    public static final String RAC_FW6_v4642 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I48&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGRC\"&\"*\")";
    public static final String STROMBO_FW1_PW3RS017_161005a = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I49&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String STROMBO_FW2_v4310 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I50&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String STROMBO_FW3_v4420 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I51&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String STROMBO_FW4_v453b = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I52&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String STROMBO_FW5_v4551 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I53&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String STROMBO_FW6_v4642 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I54&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGPC\"&\"*\")";
    public static final String DEHUM_FW1_v4310 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I55&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGAC\"&\"*\")";
    public static final String DEHUM_FW2_v4420 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I56&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGAC\"&\"*\")";
    public static final String DEHUM_FW3_v453b = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I57&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGAC\"&\"*\")";
    public static final String DEHUM_FW4_v4551 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I58&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGAC\"&\"*\")";
    public static final String DEHUM_FW5_v4642 = "COUNTIFS('" + myTools.getDate() + "'!$G:$G,$I59&\"*\", '" + myTools.getDate() + "'!$H:$H, \"FGAC\"&\"*\")";
    
    //quite a nightmarish line, I know. This adds every formula from the fw counters.
    public static final String FW_TOTAL =FRIDGE_FW1_PW1MA076 + " + " + FRIDGE_FW2_PW1MA079 + " + " + RAC_FW1_PW1RS326 + " + " + RAC_FW2_v4310 + " + " + RAC_FW3_v4420 + " + " + RAC_FW4_v453b + " + " + RAC_FW5_v4551 + " + " + RAC_FW6_v4642 + " + " + STROMBO_FW1_PW3RS017_161005a + " + " + STROMBO_FW2_v4310 + " + " + STROMBO_FW3_v4420 + " + " + STROMBO_FW4_v453b + " + " + STROMBO_FW5_v4551 + " + " + STROMBO_FW6_v4642 + " + " + DEHUM_FW1_v4310 + " + " + DEHUM_FW2_v4420 + " + " + DEHUM_FW3_v453b + " + " + DEHUM_FW4_v4551 + " + " + DEHUM_FW5_v4642;
    
    
    
    
    
    
    
//    public static final String FRIDGE_NEW = "=COUNTIFS('Jul 2'!$G:$G,\"PW1MA079*\")+COUNTIFS('Jul 2'!$G:$G,\"PW1MA076*\")";
//    public static final String FRIDGE_OLD = "=(COUNTIF('Jul 2'!$G:$G,\"PW1MA*\")+COUNTIF('Jul 2'!$G:$G,\"PW3MA*\"))-BH13";
//    public static final String RAC_X_NEW = "=COUNTIFS('Jul 2'!$G:$G,BH11,'Jul 2'!$H:$H,\"FGRC*\")+COUNTIFS('Jul 2'!$G:$G,BH11,'Jul 2'!$H:$H,\"FFRE*\")";
//    public static final String RAC_X_OLD = "=(COUNTIFS('Jul 2'!$G:$G,\"v*\",'Jul 2'!$H:$H,\"FGRC*\")+COUNTIFS('Jul 2'!$G:$G,\"PW3RS*\",'Jul 2'!$H:$H,\"FFRE*\"))-BH15";
//    public static final String RAC_GEN2_NEW = "=COUNTIFS('Jul 2'!$G:$G,\"PW1RS326\",'Jul 2'!$H:$H,\"FGRC*\")+COUNTIFS('Jul 2'!$G:$G,\"PW1RS326\",'Jul 2'!$H:$H,\"FFRE*\")";
//    public static final String RAC_GEN2_OLD = "=(COUNTIFS('Jul 2'!$G:$G,\"PW1RS*\",'Jul 2'!$H:$H,\"FFRE*\")+COUNTIFS('Jul 2'!$G:$G,\"PW1RS*\",'Jul 2'!$H:$H,\"FGRC*\"))-BH17";
//    public static final String STROMBO_NEW = "=COUNTIFS('Jul 2'!$G:$G,BH11,'Jul 2'!$H:$H,\"FGPC*\")";
//    public static final String STROMBO_OLD = "=COUNTIFS('Jul 2'!$H:$H,\"FGPC*\")-BH19";
//    public static final String DEHUM_NEW = "=COUNTIFS('Jul 2'!$G:$G,BH11,'Jul 2'!$H:$H,\"FGAC*\")";
//    public static final String DEHUM_OLD = "=COUNTIFS('Jul 2'!$H:$H,\"FGAC*\")-BH21";
    
}
