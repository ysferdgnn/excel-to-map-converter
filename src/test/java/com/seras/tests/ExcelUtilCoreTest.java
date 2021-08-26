package com.seras.tests;

import com.seras.api.ExcelUtilCore;
import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.nio.file.NoSuchFileException;

public class ExcelUtilCoreTest {

    String workingDir=System.getProperty("user.dir");
    String xssfFileName="files\\xssfTestFile.xlsx";
    String hssfFileName="files\\hssfTestFile.xls";
    String nonExcelFileName ="files\\test.txt";
    String xssfFilePath;
    String nonExcelFilePath;
    String hssfFilePath;
    File xssfFile ;
    File hssfFile;
    File nonExcelFile;

    ExcelUtilCore excelUtil;

    @Before
    public void initTest(){

        excelUtil= ExcelUtilCore.getInstance();
        xssfFilePath=workingDir+"\\"+xssfFileName;
        nonExcelFilePath=workingDir+"\\"+nonExcelFileName;
        hssfFilePath =workingDir+"\\"+hssfFileName;
        xssfFile = new File(xssfFilePath);
        nonExcelFile=new File(nonExcelFilePath);
        hssfFile=new File(hssfFilePath);
    }

    @Test
    public void testDetectIsFileValidExcelFile() throws InvalidFileExtensionNameException, NoSuchFileException, NotFileException {


        boolean isXssfFileValid= excelUtil.detectIsFileValidExcelFile(xssfFile);
        boolean isHssfFileValid = excelUtil.detectIsFileValidExcelFile(hssfFile);
        boolean wrongFileFormat = excelUtil.detectIsFileValidExcelFile(nonExcelFile);

        Assert.assertTrue(isXssfFileValid);
        Assert.assertTrue(isHssfFileValid);
        Assert.assertFalse(wrongFileFormat);
    }

    @Test
    public void testFindXmlSpreadSheetFormat() throws InValidFileFormatException, InvalidFileExtensionNameException, NoSuchFileException, NotFileException, InvalidSpreadSheetFormatException {

        File nonExcelFile=new File(nonExcelFilePath);

        SpreadSheetFormat xssfSpreadSheetFormat = excelUtil.findXmlSpreadSheetFormat(xssfFile);
        SpreadSheetFormat hssfSpreadSheetFormat = excelUtil.findXmlSpreadSheetFormat(hssfFile);


        Assert.assertEquals(xssfSpreadSheetFormat,SpreadSheetFormat.XSSF);
        Assert.assertEquals(hssfSpreadSheetFormat,SpreadSheetFormat.HSSF);
        Assert.assertThrows(InValidFileFormatException.class,()->excelUtil.findXmlSpreadSheetFormat(nonExcelFile));
    }



    @Test
    public void testReadFileFromPath() throws NoSuchFileException {
        File xssfFile = excelUtil.readFileFromPath(xssfFilePath);
        File hssfFile = excelUtil.readFileFromPath(hssfFilePath);

        Assert.assertTrue(xssfFile.exists());
        Assert.assertTrue(hssfFile.exists());
    }
}
