package com.seras.api;

import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;
import com.seras.interfaces.ExcelUtil;
import com.seras.util.FileUtils;

import java.io.File;
import java.nio.file.NoSuchFileException;

public class ExcelUtilCore implements ExcelUtil {

    /* begin singleton */
    private static  ExcelUtilCore excelUtilCore;
    private ExcelUtilCore(){};
    public static  ExcelUtilCore getInstance(){
        if (excelUtilCore==null)
            excelUtilCore=new ExcelUtilCore();
        return excelUtilCore;
    }
    /* end singleton */


    @Override
    public boolean detectIsFileValidExcelFile(File file) throws InvalidFileExtensionNameException, NoSuchFileException, NotFileException {
        String fileNameExtension= FileUtils.getInstance().getFileExtension(file);

        return fileNameExtension.toLowerCase().contentEquals("xls") || fileNameExtension.toLowerCase().contentEquals("xlsx");


    }

    @Override
    public SpreadSheetFormat findXmlSpreadSheetFormat(File file) throws InValidFileFormatException, InvalidFileExtensionNameException, NoSuchFileException, NotFileException, InvalidSpreadSheetFormatException {
        if (!detectIsFileValidExcelFile(file)){
            throw  new InValidFileFormatException("Given File has invalid file extension!");
        }

        if(FileUtils.getInstance().isFileExtensionXssfFile(file)){
            return SpreadSheetFormat.XSSF;
        }
        else if (FileUtils.getInstance().isFileExtensionHssfFile(file)){
            return SpreadSheetFormat.HSSF;
        }else {
            throw new InvalidSpreadSheetFormatException();
        }
    }

    @Override
    public File readFileFromPath(String path) throws NoSuchFileException {
        File file = new File(path);
        if (file.exists()){return  file;}
        else throw new NoSuchFileException(String.format("Can not find any file at given path! Path->%s",path));

    }
}
