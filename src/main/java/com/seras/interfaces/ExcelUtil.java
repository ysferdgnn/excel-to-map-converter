package com.seras.interfaces;

import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InValidFileFormatException;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.exceptions.NotFileException;

import java.io.File;
import java.nio.file.NoSuchFileException;

public interface ExcelUtil {

    boolean detectIsFileValidExcelFile(File file) throws InvalidFileExtensionNameException, NoSuchFileException, NotFileException;

    SpreadSheetFormat findXmlSpreadSheetFormat(File file) throws InValidFileFormatException, InvalidFileExtensionNameException, NoSuchFileException, NotFileException, InvalidSpreadSheetFormatException;

    File readFileFromPath(String path) throws NoSuchFileException;
}
