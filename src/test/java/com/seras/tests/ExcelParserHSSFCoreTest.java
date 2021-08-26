package com.seras.tests;

import com.seras.api.ExcelParserHSSFCore;
import com.seras.api.ExcelToMapConverterImplementation;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class ExcelParserHSSFCoreTest {

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

    ExcelParserHSSFCore excelParserHSSFCore ;



    @Before
    public void initTest(){

        excelParserHSSFCore=  ExcelParserHSSFCore.getInstance();
        xssfFilePath=workingDir+"\\"+xssfFileName;
        hssfFilePath=workingDir+"\\"+hssfFileName;
        nonExcelFilePath=workingDir+"\\"+nonExcelFileName;

        xssfFile = new File(xssfFilePath);
        hssfFile =new File(hssfFilePath);
        nonExcelFile=new File(nonExcelFilePath);
    }

    @Test
    public void testGetHssfWorkbookFromFile() throws IOException, InvalidSpreadSheetFormatException, InvalidFormatException {
        HSSFWorkbook workbookHssf = excelParserHSSFCore.getHssfWorkbookFromFile(hssfFile);
        Assert.assertNotNull(workbookHssf);
    }

    @Test
    public void testGetHssfSheetsFromFile() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        List<HSSFSheet> sheetsHssf = excelParserHSSFCore.getHssfSheetsFromFile(hssfFile);

        Assert.assertTrue(sheetsHssf.size()>0);

    }

    @Test
    public void  testGetHssfSheetFromFileByIndex() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFSheet sheetOverflowIndex = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,10);

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetOverflowIndex);
    }

    @Test
    public void testGetHssfSheetFromFileByName() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByName(hssfFile,"Sheet1");

        HSSFSheet sheetWrongName =excelParserHSSFCore.getHssfSheetFromFileByName(hssfFile,"Saruman");

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetWrongName);
    }

    @Test
    public void testGetHssfRowsBySheet() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        List<HSSFRow> rows = excelParserHSSFCore.getHssfRowsBySheet(sheet);

        Assert.assertTrue(rows.size()>0);
    }


    @Test
    public void testGetHssfRowBySheetAndIndex() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row= excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);
        Assert.assertNotNull(row);
    }

    @Test
    public void testGetHssfRowBySheetAndIndexOverflowIndex() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row= excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,Integer.MAX_VALUE);
        Assert.assertNull(row);
    }

    @Test
    public void  testGetHssfCellsByRow() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet=excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);
        List<HSSFCell> cells =excelParserHSSFCore.getHssfCellsByRow(row);

        Assert.assertTrue(cells.size()>0);
    }

    @Test
    public void  testGetHssfCellsByRowNullParameter()  {

        List<HSSFCell> cells =excelParserHSSFCore.getHssfCellsByRow(null);

        Assert.assertEquals(0, (int) Optional.ofNullable(cells).map(List::size).orElse(0));
    }
    @Test
    public void testGetHssfCellByRowAndIndex() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);
        HSSFCell cell=excelParserHSSFCore.getHssfCellByRowAndIndex(row,0);

        Assert.assertNotNull(cell);
    }
    

    @Test
    public void testFindCellHeaderNamesFromFirstRowHssfRowByRow() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserHSSFCore.findCellHeaderNamesFromFirstRowHssfRowByRow(row,0);
    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowHssfRowByEmptyRow() throws IOException, InvalidSpreadSheetFormatException {
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserHSSFCore.findCellHeaderNamesFromFirstRowHssfRowByRow(row,Integer.MAX_VALUE);
    }
}
