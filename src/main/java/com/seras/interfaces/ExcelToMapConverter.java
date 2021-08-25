package com.seras.interfaces;

import com.seras.enums.SpreadSheetFormat;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelToMapConverter {

   boolean detectIsFileValidExcelFile(File file);
   SpreadSheetFormat findXmlSpreadSheetFormat(File file);
   File readFileFromPath(String path);
   HSSFWorkbook getHssfWorkbookFromFile(File file);
   XSSFWorkbook getXssfWorkbookFromFile(File file);

   List<HSSFSheet>getHssfSheetsFromFile(File file);
   List<XSSFSheet>getXssfSheetsFromFile(File file);

   HSSFSheet getHssfSheetFromFileByIndex(File file,int index);
   XSSFSheet getXssfSheetFromFileByIndex(File file,int index);

   HSSFSheet getHssfSheetFromFileByName(File file,String sheet);
   XSSFSheet getXssfSheetFromFileByName(File file,String sheet);

   List<XSSFRow> getXssfRowsBySheet(XSSFSheet sheet);
   List<HSSFRow> getHssfRowsBySheet(HSSFSheet sheet);

   XSSFRow getXssfRowBySheetAndIndex(XSSFSheet sheet,int index);
   HSSFRow  getHssfRowBySheetAndIndex(HSSFSheet sheet,int index);

   List<XSSFCell> getXssfCellsByRow(XSSFRow row);
   XSSFCell getXssfCellByRowAndIndex(XSSFRow row,int index);

   List<HSSFCell> getHssfCellsByRow(HSSFRow row);
   HSSFCell getHssfCellByRowAndIndex(XSSFRow row,int index);

   List<Map<String,String>> parseExcelFileToMapList(File file);
   List<Map<String,String>> parseXssfSheetToMapList(XSSFSheet sheet);
   List<Map<String,String>> parseHssfSheetToMapList(HSSFSheet sheet);
   Map<String,String>       parseXssfRowToMap(XSSFRow row);
   Map<String,String>       parseHssfRowToMap(HSSFRow row);


}
