package com.seras.tests;

import com.seras.api.ExcelParserXSFFCore;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
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
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExcelParserXSSFCoreTest {


    String workingDir=System.getProperty("user.dir");
    String xssfFileName="files\\xssfTestFile.xlsx";

    String nonExcelFileName ="files\\test.txt";
    String xssfFilePath;
    String nonExcelFilePath;

    File xssfFile ;
    File nonExcelFile;

   ExcelParserXSFFCore excelParserXSFFCore;

    @Before
    public void initTest(){

        excelParserXSFFCore= ExcelParserXSFFCore.getInstance();
        xssfFilePath=workingDir+"\\"+xssfFileName;
        nonExcelFilePath=workingDir+"\\"+nonExcelFileName;
        xssfFile = new File(xssfFilePath);
        nonExcelFile=new File(nonExcelFilePath);
    }

    @Test
    public void testParseXssfRowToMap() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet =excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row= excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);

        Map<String,String> map = excelParserXSFFCore.parseXssfRowToMap(row,false,0);


    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowXssfRowByRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet =excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserXSFFCore.findCellHeaderNamesFromFirstRowXssfRowByRow(row,0 );
        Assert.assertNotEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowXssfRowByEmptyRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet =excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserXSFFCore.findCellHeaderNamesFromFirstRowXssfRowByRow(row,Integer.MAX_VALUE );
        Assert.assertEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public  void testGetXssfCellByRowAndIndexOverflowIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet, 0);
        XSSFCell cell = excelParserXSFFCore.getXssfCellByRowAndIndex(row,Integer.MAX_VALUE);
        Assert.assertNull(cell);

    }

    @Test
    public  void testGetXssfCellsByRowNullParameter() {
        List<XSSFCell> cells = excelParserXSFFCore.getXssfCellsByRow(null);

        Assert.assertEquals(0, (cells == null ? 0 : cells.size()));
    }
    @Test
    public  void testGetXssfCellByRowAndIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet, 0);
        XSSFCell cell = excelParserXSFFCore.getXssfCellByRowAndIndex(row,0);
        Assert.assertNotNull(cell);

    }

    @Test
    public  void testGetXssfCellsByRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet=excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row= excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        List<XSSFCell> cells = excelParserXSFFCore.getXssfCellsByRow(row);

        Assert.assertTrue(cells.size()>0);
    }

    @Test
    public  void testGetXssfRowBySheetAndIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet= excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        Assert.assertNotNull(row);
    }

    @Test
    public  void testGetXssfRowBySheetAndIndexOverflowIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet= excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,Integer.MAX_VALUE);
        Assert.assertNull(row);
    }

    @Test
    public  void testGetXssfSheetFromFileByName() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByName(xssfFile,"Sheet1");
        XSSFSheet sheetWrong = excelParserXSFFCore.getXssfSheetFromFileByName(xssfFile,"SAURON");

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetWrong);
    }
    @Test
    public void testGetXssfRowsBySheet() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        List<XSSFRow> rows = excelParserXSFFCore.getXssfRowsBySheet(sheet);

        Assert.assertTrue(rows.size()>0);
    }

    @Test
    public void  testGetXssfSheetFromFileByIndex() throws IOException, InvalidSpreadSheetFormatException, InvalidFormatException {
        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFSheet sheetOverflowIndex = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,10);

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetOverflowIndex);
    }

    @Test
    public void testGetXssfSheetsFromFile() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        List<XSSFSheet> sheetsXssf = excelParserXSFFCore.getXssfSheetsFromFile(xssfFile);

        Assert.assertTrue(sheetsXssf.size()>0);

    }

    @Test
    public void testGetXssfWorkbookFromFile() throws  IOException, InvalidSpreadSheetFormatException, InvalidFormatException {
        XSSFWorkbook workbookXssf = excelParserXSFFCore.getXssfWorkbookFromFile(xssfFile);
        Assert.assertNotNull(workbookXssf);
    }


    
}
