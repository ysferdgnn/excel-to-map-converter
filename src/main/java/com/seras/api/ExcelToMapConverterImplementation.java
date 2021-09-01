package com.seras.api;

import com.seras.constants.LogConstants;
import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import com.seras.interfaces.ExcelToMapConverter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;
import java.util.stream.Collectors;

public class ExcelToMapConverterImplementation implements ExcelToMapConverter {

    Logger logger = Logger.getLogger(ExcelToMapConverterImplementation.class.getName());

    private static  int rowPointer;
    private static  int sheetPointer;
    private static boolean isFirstRowCellHeader;

    private static  ExcelToMapConverterImplementation excelToMapConverterImplementation;
    private ExcelToMapConverterImplementation(){}
    public static  ExcelToMapConverterImplementation getInstance(){
        if (excelToMapConverterImplementation==null)
            excelToMapConverterImplementation=new ExcelToMapConverterImplementation();

        rowPointer=0;
        sheetPointer=0;
        isFirstRowCellHeader=false;
        return  excelToMapConverterImplementation;
    }


    @Override
    public List<Map<String, String>> parseExcelFileToMapList(File file) throws InValidFileFormatException {

        if(file==null){
            logger.warning(LogConstants.getInstance().msgFileEmpty);
            return  null;
        }
        if(rowPointer<0){
            logger.warning("invalid row pointer!");
            return  null;
        }

        try {
            SpreadSheetFormat spreadSheetFormat = ExcelUtilCore.getInstance().findXmlSpreadSheetFormat(file);

            if (spreadSheetFormat==SpreadSheetFormat.HSSF){
                return ExcelParserHSSFCore
                        .getInstance()
                        .setSheetPointer(sheetPointer)
                        .setRowPointer(rowPointer)
                        .setIsFirstRowCellHeader(isFirstRowCellHeader)
                        .parseHssfFileToMapList(file);
            }
            else if (spreadSheetFormat==SpreadSheetFormat.XSSF)
            {
                return ExcelParserXSFFCore
                        .getInstance()
                        .setSheetPointer(sheetPointer)
                        .setRowPointer(rowPointer)
                        .setIsFirstRowCellHeader(isFirstRowCellHeader)
                        .parseXssfFileToMapList(file);
            }
            else{

               throw new InvalidSpreadSheetFormatException();
            }

        } catch ( InvalidFileExtensionNameException | NotFileException | InvalidSpreadSheetFormatException | IOException | InvalidFormatException e) {
            e.printStackTrace();
            return null;
        }
        catch (InValidFileFormatException e){
            throw e;
        }
    }

    @Override
    public List<String> findHeadersFromMapList(List<Map<String, String>> mapList) {
        if (mapList==null)
            return null;
        if (mapList.get(0)==null)
            return null;

        Map<String,String> map = mapList.get(0);
        return new ArrayList<>(map.keySet());
    }
    public ExcelToMapConverterImplementation setRowPointer(int _rowPointer){
        ExcelToMapConverterImplementation.rowPointer=_rowPointer;
        return  this;
    }
    public ExcelToMapConverterImplementation setSheetPointer(int _sheetPointer){
        ExcelToMapConverterImplementation.sheetPointer=_sheetPointer;
        return this;
    }
    public ExcelToMapConverterImplementation setIsFirstRowCellHeader(boolean _isIsFirstRowCellHeader){
        ExcelToMapConverterImplementation.isFirstRowCellHeader=_isIsFirstRowCellHeader;
        return this;
    }

}
