package com.seras.constants;

public class FileNameConstants {

    private static  FileNameConstants fileNameConstants;

    private FileNameConstants(){};

    public static FileNameConstants  getInstance(){
        if (fileNameConstants==null){
            fileNameConstants=new FileNameConstants();
        }
        return fileNameConstants;
    }

    public String xsffFileExtension="xlsx";
    public String hssfFileExtension="xls";

}
