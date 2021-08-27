package com.seras.tests;

import com.seras.api.ExcelParserHSSFCore;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.testConstants.TestFileConstants;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExcelParserHSSFCoreTest {



    ExcelParserHSSFCore excelParserHSSFCore ;



    @Before
    public void initTest(){

        excelParserHSSFCore=  ExcelParserHSSFCore.getInstance();
    }

    @Test
    public void testGetHssfWorkbookFromFile() throws IOException, InvalidSpreadSheetFormatException {
        HSSFWorkbook workbookHssf = excelParserHSSFCore.getHssfWorkbookFromFile(TestFileConstants.getInstance().getHssfFile());
        Assert.assertNotNull(workbookHssf);
    }

    @Test
    public void testGetHssfSheetsFromFile() throws IOException, InvalidSpreadSheetFormatException {
        List<HSSFSheet> sheetsHssf = excelParserHSSFCore.getHssfSheetsFromFile(TestFileConstants.getInstance().getHssfFile());

        Assert.assertTrue(sheetsHssf.size()>0);

    }

    @Test
    public void  testGetHssfSheetFromFileByIndex() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFSheet sheetOverflowIndex = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,10);

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetOverflowIndex);
    }

    @Test
    public void testGetHssfSheetFromFileByName() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByName(hssfFile,"Sheet1");

        HSSFSheet sheetWrongName =excelParserHSSFCore.getHssfSheetFromFileByName(hssfFile,"Saruman");

        Assert.assertNotNull(sheet);
        Assert.assertNull(sheetWrongName);
    }

    @Test
    public void testGetHssfRowsBySheet() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        List<HSSFRow> rows = excelParserHSSFCore.getHssfRowsBySheet(sheet);

        Assert.assertTrue(rows.size()>0);
    }


    @Test
    public void testGetHssfRowBySheetAndIndex() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row= excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);
        Assert.assertNotNull(row);
    }

    @Test
    public void testGetHssfRowBySheetAndIndexOverflowIndex() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet =excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row= excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,Integer.MAX_VALUE);
        Assert.assertNull(row);
    }

    @Test
    public void  testGetHssfCellsByRow() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
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
        File hssfFile =TestFileConstants.getInstance().getHssfFile();
        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);
        HSSFCell cell=excelParserHSSFCore.getHssfCellByRowAndIndex(row,0);

        Assert.assertNotNull(cell);
    }
    

    @Test
    public void testFindCellHeaderNamesFromFirstRowHssfRowByRow() throws IOException, InvalidSpreadSheetFormatException {
        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserHSSFCore.findCellHeaderNamesFromFirstRowHssfRowByRow(row);
        Assert.assertNotEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public void testFindCellHeaderNamesFromFirstRowHssfRowByEmptyRow() throws IOException, InvalidSpreadSheetFormatException {

        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserHSSFCore.setRowPointer(Integer.MAX_VALUE).findCellHeaderNamesFromFirstRowHssfRowByRow(row);
        Assert.assertEquals(Optional.of(0).get(),Optional.ofNullable(headerList).map(List::size).orElse(0));
    }

    @Test
    public void testFindCellHeaderNamesByFirstHssfRowDefault() throws IOException, InvalidSpreadSheetFormatException {

        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        HSSFRow row = excelParserHSSFCore.getHssfRowBySheetAndIndex(sheet,0);

        List<String> headerList = excelParserHSSFCore.findCellHeaderNamesByFirstHssfRowDefault(row);

        assert headerList !=null;

        Assert.assertNotEquals(0, headerList.size());
        Assert.assertEquals("A",headerList.get(0));
    }

    @Test
    public void  testParseHssfSheetToMapList() throws IOException, InvalidSpreadSheetFormatException {

        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        List<Map<String,String>> maps = excelParserHSSFCore.parseHssfSheetToMapList(sheet);

        assert maps !=null;

        Assert.assertNotEquals(0,maps.size());
    }

    @Test
    public void  testParseHssfSheetToMapListWithFirstRowAsHeader() throws IOException, InvalidSpreadSheetFormatException {

        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFSheet sheet = excelParserHSSFCore.getHssfSheetFromFileByIndex(hssfFile,0);
        List<Map<String,String>> maps = excelParserHSSFCore.setIsFirstRowCellHeader(true).parseHssfSheetToMapList(sheet);

        assert maps !=null;

        Assert.assertNotEquals(0,maps.size());
        Assert.assertTrue(maps.stream().findFirst().get().keySet().stream().anyMatch(s->s.contentEquals("ad")));
    }

    @Test
    public  void testParseHssfWorkBookToMapList() throws IOException, InvalidSpreadSheetFormatException {

        File hssfFile =TestFileConstants.getInstance().getHssfFile();

        HSSFWorkbook workbook = excelParserHSSFCore.getHssfWorkbookFromFile(hssfFile);
        List<Map<String,String>> maps = excelParserHSSFCore.parseHssfWorkBookToMapList(workbook);

        assert maps!=null;
        Assert.assertNotEquals(0,maps.size());
    }
}
