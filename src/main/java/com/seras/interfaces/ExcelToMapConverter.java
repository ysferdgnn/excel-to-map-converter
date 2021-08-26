package com.seras.interfaces;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelToMapConverter {




   List<Map<String,String>> parseExcelFileToMapList(File file,boolean isFirstRowCellHeader,int startPos);




}
