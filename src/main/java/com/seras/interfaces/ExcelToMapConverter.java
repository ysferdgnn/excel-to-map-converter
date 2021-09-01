package com.seras.interfaces;

import com.seras.exceptions.InValidFileFormatException;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelToMapConverter {




   List<Map<String,String>> parseExcelFileToMapList(File file) throws InValidFileFormatException;
   List<String> findHeadersFromMapList(List<Map<String,String>> mapList);




}
