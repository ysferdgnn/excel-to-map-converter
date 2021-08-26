package com.seras.interfaces;

import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface ExcelParserXSSF {

    XSSFWorkbook getXssfWorkbookFromFile(File file) throws InValidFileFormatException, InvalidFileExtensionNameException, IOException, NotFileException, InvalidSpreadSheetFormatException, InvalidFormatException;

    List<XSSFSheet> getXssfSheetsFromFile(File file) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException;

    XSSFSheet getXssfSheetFromFileByIndex(File file,int index) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException;

    XSSFSheet getXssfSheetFromFileByName(File file,String sheet) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException;

    List<XSSFRow> getXssfRowsBySheet(XSSFSheet sheet);

    XSSFRow getXssfRowBySheetAndIndex(XSSFSheet sheet,int index);

    List<XSSFCell> getXssfCellsByRow(XSSFRow row);

    XSSFCell getXssfCellByRowAndIndex(XSSFRow row,int index);

    List<Map<String,String>> parseXssfSheetToMapList(XSSFSheet sheet, boolean isFirstRowCellHeader, int startPos);

    Map<String,String>       parseXssfRowToMap(XSSFRow row,boolean isFirstRowCellHeader,int startPos);

    List<String> findCellHeaderNamesFromFirstRowXssfRowByRow(XSSFRow row,int startPos);

    List<String > findCellHeaderNamesByFirstXssfRowDefault(XSSFRow row);
}
