package com.seras.interfaces;

import com.seras.exceptions.InValidFileFormatException;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelToMapConverter {




   List<Map<String,String>> parseExcelFileToMapList(File file) throws InValidFileFormatException;
   List<String> findHeadersFromMapList(List<Map<String,String>> mapList);
   List<String> findSheetNamesFromFile(File file) throws InValidFileFormatException;
   Integer findRowCountFromSheetName(File file,String sheetName) throws InValidFileFormatException;
   Integer findSheetNumberFromSheetName(File file,String sheetName) throws InValidFileFormatException;




}
