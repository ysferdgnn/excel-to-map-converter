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
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class ExcelToMapConverterImplementation implements ExcelToMapConverter {

    Logger logger = Logger.getLogger(ExcelToMapConverterImplementation.class.getName());


    @Override
    public List<Map<String, String>> parseExcelFileToMapList(File file, boolean isFirstRowCellHeader, int rowPointer,int sheetPointer) {

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



        } catch (InValidFileFormatException | InvalidFileExtensionNameException | NotFileException | InvalidSpreadSheetFormatException | IOException | InvalidFormatException e) {
            e.printStackTrace();
            return null;
        }


    }
}
