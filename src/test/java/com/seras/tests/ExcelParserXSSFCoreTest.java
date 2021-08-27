package com.seras.tests;

import com.seras.api.ExcelParserXSFFCore;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.testConstants.TestFileConstants;
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




   ExcelParserXSFFCore excelParserXSFFCore;

    @Before
    public void initTest(){

        excelParserXSFFCore= ExcelParserXSFFCore.getInstance();

    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowXssfRowByRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet =excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserXSFFCore.findCellHeaderNamesFromFirstRowXssfRowByRow(row);
        Assert.assertNotEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowXssfRowByEmptyRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet =excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserXSFFCore.setRowPointer(Integer.MAX_VALUE).findCellHeaderNamesFromFirstRowXssfRowByRow(row );
        Assert.assertEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public  void testGetXssfCellByRowAndIndexOverflowIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

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

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet, 0);
        XSSFCell cell = excelParserXSFFCore.getXssfCellByRowAndIndex(row,0);
        Assert.assertNotNull(cell);

    }

    @Test
    public  void testGetXssfCellsByRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet=excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row= excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        List<XSSFCell> cells = excelParserXSFFCore.getXssfCellsByRow(row);

        Assert.assertTrue(cells.size()>0);
    }

    @Test
    public  void testGetXssfRowBySheetAndIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet= excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        Assert.assertNotNull(row);
    }

    @Test
    public  void testGetXssfRowBySheetAndIndexOverflowIndex() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet= excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,Integer.MAX_VALUE);
        Assert.assertNull(row);
    }

    @Test
    public  void testGetXssfSheetFromFileByName() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByName(xssfFile,"Sheet1");
        XSSFSheet sheetWrong = excelParserXSFFCore.getXssfSheetFromFileByName(xssfFile,"SAURON");

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetWrong);
    }
    @Test
    public void testGetXssfRowsBySheet() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        List<XSSFRow> rows = excelParserXSFFCore.getXssfRowsBySheet(sheet);

        Assert.assertTrue(rows.size()>0);
    }

    @Test
    public void  testGetXssfSheetFromFileByIndex() throws IOException, InvalidSpreadSheetFormatException, InvalidFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFSheet sheetOverflowIndex = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,10);

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetOverflowIndex);
    }

    @Test
    public void testGetXssfSheetsFromFile() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        List<XSSFSheet> sheetsXssf = excelParserXSFFCore.getXssfSheetsFromFile(xssfFile);

        Assert.assertTrue(sheetsXssf.size()>0);

    }

    @Test
    public void testGetXssfWorkbookFromFile() throws  IOException, InvalidSpreadSheetFormatException, InvalidFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFWorkbook workbookXssf = excelParserXSFFCore.getXssfWorkbookFromFile(xssfFile);
        Assert.assertNotNull(workbookXssf);
    }

    @Test
    public void testParseXssfRowToMap() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        Map<String,String> map = excelParserXSFFCore.parseXssfRowToMap(row);
        Assert.assertNotEquals(Optional.of(0).get(),Optional.ofNullable(map).map(Map::size).orElse(0));

    }

    @Test
    public void testParseXssfRowToMapFirstRowAsCellHeader() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,0);
        Map<String,String> map = excelParserXSFFCore.setIsFirstRowCellHeader(true).parseXssfRowToMap(row);
        Assert.assertNotEquals(Optional.of(0).get(),Optional.ofNullable(map).map(Map::size).orElse(0));

        assert map != null;
        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("ad")));
        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("soyad")));
        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("tc")));

        Assert.assertEquals(map.get("ad"),"ad" );
        Assert.assertEquals(map.get("soyad"),"soyad" );
        Assert.assertEquals(map.get("tc"),"tc" );

    }

    @Test
    public void testParseXssfRowToMapFirstRowAsCellHeaderSecondRow() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        XSSFRow row = excelParserXSFFCore.getXssfRowBySheetAndIndex(sheet,1);
        Map<String,String> map = excelParserXSFFCore.setIsFirstRowCellHeader(true).parseXssfRowToMap(row);

        assert map !=null;
        Assert.assertNotEquals(0,map.size());

        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("ad")));
        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("soyad")));
        Assert.assertTrue(map.keySet().stream().anyMatch(s->s.contentEquals("tc")));

        Assert.assertEquals(map.get("ad"),"yusuf" );
        Assert.assertEquals(map.get("soyad"),"erdoğan" );
        Assert.assertEquals(map.get("tc"),"231.0" );



    }

    @Test
    public void testParseXssfSheetToMapList() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(xssfFile,0);
        List<Map<String,String>> maps = excelParserXSFFCore.parseXssfSheetToMapList(sheet);

        assert maps !=null;

        Assert.assertNotEquals(0,maps.size());
        Assert.assertEquals(7,maps.size());
    }

    @Test
    public void testParseXssfSheetToMapListAsyncColumns() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File asyncColumnsExcelFile = TestFileConstants.getInstance().getXssfFileAsyncCols();

        XSSFSheet sheet = excelParserXSFFCore.getXssfSheetFromFileByIndex(asyncColumnsExcelFile,0);
        List<Map<String,String>> maps = excelParserXSFFCore.setIsFirstRowCellHeader(true).parseXssfSheetToMapList(sheet);

        assert maps !=null;

        Assert.assertNotEquals(0,maps.size());

    }

    @Test
    public void testParseXssfWorkBookToMapList() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File xssfFile = TestFileConstants.getInstance().getXssfFile();

       XSSFWorkbook workbook = excelParserXSFFCore.getXssfWorkbookFromFile(xssfFile);
       List<Map<String,String>> maps = excelParserXSFFCore.parseXssfWorkBookToMapList(workbook);

       assert maps !=null;

       Assert.assertNotEquals(0, maps.size());

    }

    @Test
    public void testParseXssfWorkBookToMapListAsyncRowExcelFile() throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {

        File asyncColumnsExcelFile = TestFileConstants.getInstance().getXssfFileAsyncCols();

        XSSFWorkbook workbook = excelParserXSFFCore.getXssfWorkbookFromFile(asyncColumnsExcelFile);
        List<Map<String,String>> maps = excelParserXSFFCore.parseXssfWorkBookToMapList(workbook);

        assert maps !=null;

        Assert.assertNotEquals(0, maps.size());

    }

    @Test
    public void testParseXssfWorkBookToMapListAsyncRowExcelFileWithFirstRowsAsHeader() throws IOException, InvalidFormatException,
            InvalidSpreadSheetFormatException{

        File asyncColumnsExcelFile = TestFileConstants.getInstance().getXssfFileAsyncCols();


        XSSFWorkbook workbook = excelParserXSFFCore.setIsFirstRowCellHeader(true).getXssfWorkbookFromFile(asyncColumnsExcelFile);
        List<Map<String,String>> maps = excelParserXSFFCore.parseXssfWorkBookToMapList(workbook);

        assert maps !=null;

        Assert.assertNotEquals(0, maps.size());

    }

    
}
