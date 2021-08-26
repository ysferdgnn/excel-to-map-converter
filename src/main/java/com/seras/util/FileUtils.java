package com.seras.util;

import com.seras.constants.FileNameConstants;
import com.seras.exceptions.InvalidFileExtensionNameException;
import com.seras.exceptions.NotFileException;

import java.io.File;
import java.nio.file.NoSuchFileException;
import java.util.Optional;

public class FileUtils {

    private static  FileUtils fileUtils;
    private FileUtils(){};

    public static FileUtils getInstance(){
        if (fileUtils==null){
            fileUtils=new FileUtils();
        }
        return  fileUtils;
    }


    public String getFileExtension(File file) throws NoSuchFileException, NotFileException, InvalidFileExtensionNameException {

        if (!file.exists()) throw new NoSuchFileException("Can not find any file at given path ->"+file.getAbsolutePath());
        if (!file.isFile()) throw  new NotFileException();


        String fileName = file.getName();
        String fileNameExtension= Optional.of(fileName)
                .filter(s->s.contains("."))
                .map(s->s.substring(fileName.lastIndexOf(".")+1))
                .orElse(null);

        if (fileNameExtension==null){
            throw  new InvalidFileExtensionNameException();
        }

        return  fileNameExtension.toLowerCase();
    }
    public boolean isFileExtensionXssfFile(File file){
        try {
            String fileExtension = getFileExtension(file);
            return fileExtension.contentEquals(FileNameConstants.getInstance().xsffFileExtension);
        }catch (Exception e){
            e.printStackTrace();
            return false;
        }
    }
    public  boolean isFileExtensionHssfFile(File file){
        try{
            String fileExtension = getFileExtension(file);
            return fileExtension.contentEquals(FileNameConstants.getInstance().hssfFileExtension);
        }catch (Exception e){
            e.printStackTrace();
            return false;
        }
    }

}
