package com.seras.tests;

import com.seras.api.ExcelToMapConverterImplementation;
import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.NoSuchFileException;
import java.util.List;
import java.util.Map;
import java.util.Optional;

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
