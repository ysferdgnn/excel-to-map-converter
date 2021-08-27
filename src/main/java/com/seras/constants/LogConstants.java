package com.seras.constants;

public class LogConstants {

    private static LogConstants logConstants;
    private LogConstants(){};
    public static  LogConstants getInstance(){
        if (logConstants==null)
            logConstants=new LogConstants();
        return logConstants;
    }

    public String msgRowEmpty ="Row is empty. Return null";
    public String msgHeaderListEmpty ="HeaderList is empty. Return null";
    public String msgRowListEmpty ="Row list is empty. Return null";
    public String msgSheetEmty="Sheet is empty.Return null";
    public String msgWorkbookEmpty ="Workbook is empty. Return null";
    public String msgFileEmpty ="File is empty.Return null";
}
