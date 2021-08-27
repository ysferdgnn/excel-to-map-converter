package com.seras.tests;

import com.seras.api.ExcelToMapConverterImplementation;
import org.junit.Before;

import java.io.File;

public class ExcelToMapConverterTest {



    String workingDir=System.getProperty("user.dir");
    String xssfFileName="files\\xssfTestFile.xlsx";
    String hssfFileName="files\\hssfTestFile.xls";
    String nonExcelFileName ="files\\test.txt";

    String xssfFilePath;
    String hssfFilePath;
    String nonExcelFilePath;

    File xssfFile ;
    File hssfFile ;
    File nonExcelFile;

    ExcelToMapConverterImplementation excelToMapConverterImplementation ;



    @Before
    public void initTest(){

        excelToMapConverterImplementation= new ExcelToMapConverterImplementation();
        xssfFilePath=workingDir+"\\"+xssfFileName;
        hssfFilePath=workingDir+"\\"+hssfFileName;
        nonExcelFilePath=workingDir+"\\"+nonExcelFileName;

         xssfFile = new File(xssfFilePath);
         hssfFile =new File(hssfFilePath);
         nonExcelFile=new File(nonExcelFilePath);
    }







}
