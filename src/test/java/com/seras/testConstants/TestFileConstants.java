package com.seras.testConstants;

import java.io.File;

public class TestFileConstants {

    private static  TestFileConstants testFileConstants;
    private  TestFileConstants(){}
    public  static TestFileConstants getInstance(){
        if (testFileConstants==null)
            testFileConstants=new TestFileConstants();
        return testFileConstants;
    }

  private final  String xssfFileName="\\files\\xssfTestFile.xlsx";
  private final  String hssfFileName="\\files\\hssfTestFile.xls";
  private final  String nonExcelFileName ="\\files\\test.txt";
  private final  String hssfFileAsyncColumnsName="\\files\\hssfTestFileAsyncColumns.xls";
  private final  String xssfFileAsyncColumnsName="\\files\\xssfTestFileAsyncColumns.xlsx";
  private final  String workingDir=System.getProperty("user.dir");

  private final String xssfFilePath=workingDir+xssfFileName;
  private final String hssfFilePath=workingDir+hssfFileName;
  private final String nonExcelFilePath=workingDir+nonExcelFileName;
  private final String hssfFileAsyncColumnsPath=workingDir+hssfFileAsyncColumnsName;
  private final String xssfFileAsyncColumnsPath=workingDir+xssfFileAsyncColumnsName;

  private File xssfFile;
  private File hssfFile;
  private File nonExcelFile;
  private File hssfFileAsyncCols;
  private File xssfFileAsyncCols;

  public File getXssfFile(){
      xssfFile=new File(xssfFilePath);
      return xssfFile;
  }

  public File getHssfFile(){
      hssfFile=new File(hssfFilePath);
      return hssfFile;
  }
  public File getNonExcelFile(){
      nonExcelFile=new File(nonExcelFilePath);
      return nonExcelFile;
  }

  public File getHssfFileAsyncCols(){
      hssfFileAsyncCols=new File(hssfFileAsyncColumnsPath);
      return hssfFileAsyncCols;
  }
  public File getXssfFileAsyncCols(){
      xssfFileAsyncCols = new File(xssfFileAsyncColumnsPath);
      return xssfFileAsyncCols;
  }



}
