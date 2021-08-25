package com.seras.api;

import com.seras.enums.SpreadSheetFormat;
import com.seras.interfaces.ExcelToMapConverter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.List;
import java.util.Map;

public class ExcelToMapConverterImplementation implements ExcelToMapConverter {

    @Override
    public boolean detectIsFileValidExcelFile(File file) {
        return false;
    }

    @Override
    public SpreadSheetFormat findXmlSpreadSheetFormat(File file) {
        return null;
    }

    @Override
    public File readFileFromPath(String path) {
        return null;
    }

    @Override
    public HSSFWorkbook getHssfWorkbookFromFile(File file) {
        return null;
    }

    @Override
    public XSSFWorkbook getXssfWorkbookFromFile(File file) {
        return null;
    }

    @Override
    public List<HSSFSheet> getHssfSheetsFromFile(File file) {
        return null;
    }

    @Override
    public List<XSSFSheet> getXssfSheetsFromFile(File file) {
        return null;
    }

    @Override
    public HSSFSheet getHssfSheetFromFileByIndex(File file, int index) {
        return null;
    }

    @Override
    public XSSFSheet getXssfSheetFromFileByIndex(File file, int index) {
        return null;
    }

    @Override
    public HSSFSheet getHssfSheetFromFileByName(File file, String sheet) {
        return null;
    }

    @Override
    public XSSFSheet getXssfSheetFromFileByName(File file, String sheet) {
        return null;
    }

    @Override
    public List<XSSFRow> getXssfRowsBySheet(XSSFSheet sheet) {
        return null;
    }

    @Override
    public List<HSSFRow> getHssfRowsBySheet(HSSFSheet sheet) {
        return null;
    }

    @Override
    public XSSFRow getXssfRowBySheetAndIndex(XSSFSheet sheet, int index) {
        return null;
    }

    @Override
    public HSSFRow getHssfRowBySheetAndIndex(HSSFSheet sheet, int index) {
        return null;
    }

    @Override
    public List<XSSFCell> getXssfCellsByRow(XSSFRow row) {
        return null;
    }

    @Override
    public XSSFCell getXssfCellByRowAndIndex(XSSFRow row, int index) {
        return null;
    }

    @Override
    public List<HSSFCell> getHssfCellsByRow(HSSFRow row) {
        return null;
    }

    @Override
    public HSSFCell getHssfCellByRowAndIndex(XSSFRow row, int index) {
        return null;
    }

    @Override
    public List<Map<String, String>> parseExcelFileToMapList(File file) {
        return null;
    }

    @Override
    public List<Map<String, String>> parseXssfSheetToMapList(XSSFSheet sheet) {
        return null;
    }

    @Override
    public List<Map<String, String>> parseHssfSheetToMapList(HSSFSheet sheet) {
        return null;
    }

    @Override
    public Map<String, String> parseXssfRowToMap(XSSFRow row) {
        return null;
    }

    @Override
    public Map<String, String> parseHssfRowToMap(HSSFRow row) {
        return null;
    }
}
