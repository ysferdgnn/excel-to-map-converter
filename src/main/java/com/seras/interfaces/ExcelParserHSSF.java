package com.seras.interfaces;

import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface ExcelParserHSSF {

    HSSFWorkbook getHssfWorkbookFromFile(File file) throws InValidFileFormatException, InvalidFileExtensionNameException,
            IOException, NotFileException, InvalidSpreadSheetFormatException;

    List<HSSFSheet> getHssfSheetsFromFile(File file) throws IOException, InvalidSpreadSheetFormatException;

    HSSFSheet getHssfSheetFromFileByIndex(File file,int index) throws IOException, InvalidSpreadSheetFormatException;

    HSSFSheet getHssfSheetFromFileByName(File file,String sheet) throws IOException, InvalidSpreadSheetFormatException;

    List<HSSFRow> getHssfRowsBySheet(HSSFSheet sheet);

    HSSFRow  getHssfRowBySheetAndIndex(HSSFSheet sheet,int index);

    List<HSSFCell> getHssfCellsByRow(HSSFRow row);

    HSSFCell getHssfCellByRowAndIndex(HSSFRow row,int index);

    List<Map<String,String>> parseHssfSheetToMapList(HSSFSheet sheet);

    Map<String,String> parseHssfRowToMap(HSSFRow row);

    List<Map<String,String>> parseHssfWorkBookToMapList(HSSFWorkbook workbook);

    List<Map<String,String>> parseHssfFileToMapList(File file) throws IOException, InvalidSpreadSheetFormatException;

    List<String> findCellHeaderNamesFromFirstRowHssfRowByRow(HSSFRow row);

    List<String> findCellHeaderNamesByFirstHssfRowDefault(HSSFRow row);


}
