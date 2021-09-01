package com.seras.tests;

import com.seras.api.ExcelToMapConverterImplementation;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.testConstants.TestFileConstants;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class ExcelToMapConverterImplementationTest {


    ExcelToMapConverterImplementation excelToMapConverterImplementation ;



    @Before
    public void initTest(){

        excelToMapConverterImplementation= ExcelToMapConverterImplementation.getInstance();

    }


    @Test
    public void testXssfFileToMap() throws InValidFileFormatException {
        File file = TestFileConstants.getInstance().getXssfFile();
       List<Map<String,String>>mapList =
               excelToMapConverterImplementation.parseExcelFileToMapList(file);

       assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
    }

    @Test
    public void testXssfFileToMapSecondRow() throws InValidFileFormatException {
        int rowPointer=1;
        File file = TestFileConstants.getInstance().getXssfFile();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setRowPointer(rowPointer).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(rowPointer);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("yusuf",mapFirst.get("A"));
    }

    @Test
    public void testXssfFileToMapFirstRowHeader() throws InValidFileFormatException {

        File file = TestFileConstants.getInstance().getXssfFile();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setIsFirstRowCellHeader(true).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("ad")));
        Assert.assertEquals("ad",mapFirst.get("ad"));
    }

    @Test
    public void testHssfFileToMap() throws InValidFileFormatException {
        File file = TestFileConstants.getInstance().getHssfFile();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
    }

    @Test
    public void testHssfFileToMapSecondRow() throws InValidFileFormatException {
        int rowPointer=1;
        File file = TestFileConstants.getInstance().getHssfFile();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setRowPointer(rowPointer).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(rowPointer);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("yusuf",mapFirst.get("A"));
    }

    @Test
    public void testHssfFileToMapFirstRowHeader() throws InValidFileFormatException {

        File file = TestFileConstants.getInstance().getHssfFile();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setIsFirstRowCellHeader(true).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("ad")));
        Assert.assertEquals("ad",mapFirst.get("ad"));
    }

    @Test
    public void testNotExcelFileToMap()  {
        File file = TestFileConstants.getInstance().getNonExcelFile();


        Assert.assertThrows(
                InValidFileFormatException.class,
                ()-> excelToMapConverterImplementation.parseExcelFileToMapList(file)
        );
    }

    @Test
    public void testXssfAsyncFileToMap() throws InValidFileFormatException {
        File file = TestFileConstants.getInstance().getXssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("ad",mapFirst.get("A"));
    }

    @Test
    public void testXssfAsyncFileToMapFirstRowHeader() throws InValidFileFormatException {
        File file = TestFileConstants.getInstance().getXssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setIsFirstRowCellHeader(true).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("ad")));
        Assert.assertEquals("ad",mapFirst.get("ad"));
    }

    @Test
    public void testXssfAsyncFileToMapSecondRow() throws InValidFileFormatException {
        int rowPointer=1;
        File file = TestFileConstants.getInstance().getXssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setRowPointer(rowPointer).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(rowPointer);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("yusuf",mapFirst.get("A"));
    }

    @Test
    public void testHssfAsyncFileToMap() throws InValidFileFormatException {
        File file = TestFileConstants.getInstance().getHssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("ad",mapFirst.get("A"));
    }

    @Test
    public void testHssfAsyncFileToMapSecondRow() throws InValidFileFormatException {
        int rowPointer=1;
        File file = TestFileConstants.getInstance().getHssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setRowPointer(rowPointer).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(rowPointer);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("A")));
        Assert.assertEquals("yusuf",mapFirst.get("A"));
    }

    @Test
    public void testHssfAsyncFileToMapFirstRowHeader() throws InValidFileFormatException {

        File file = TestFileConstants.getInstance().getHssfFileAsyncCols();
        List<Map<String,String>>mapList =
                excelToMapConverterImplementation.setIsFirstRowCellHeader(true).parseExcelFileToMapList(file);

        assert mapList !=null;
        Assert.assertNotEquals(0,mapList.size());

        Map<String,String> mapFirst = mapList.get(0);

        Assert.assertTrue( mapFirst.keySet().stream().anyMatch(s-> s.contentEquals("ad")));
        Assert.assertEquals("ad",mapFirst.get("ad"));
    }









}
