package com.seras.api;

import com.seras.interfaces.ExcelToMapConverter;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class ExcelToMapConverterImplementation implements ExcelToMapConverter {

    Logger logger = Logger.getLogger(ExcelToMapConverterImplementation.class.getName());


    @Override
    public List<Map<String, String>> parseExcelFileToMapList(File file, boolean isFirstRowCellHeader, int startPos) {
        return null;
    }
}
