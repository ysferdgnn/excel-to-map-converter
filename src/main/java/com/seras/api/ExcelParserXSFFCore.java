package com.seras.api;

import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.interfaces.ExcelParserXSSF;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.logging.Logger;

public class ExcelParserXSFFCore implements ExcelParserXSSF {

    /* begin singleton */
    private static  ExcelParserXSFFCore excelParserXSFFCore;
    private ExcelParserXSFFCore(){};
    public static ExcelParserXSFFCore getInstance(){
        if (excelParserXSFFCore==null)
            excelParserXSFFCore=new ExcelParserXSFFCore();
        return excelParserXSFFCore;
    }
    /* end singleton */

    Logger logger = Logger.getLogger(ExcelParserXSFFCore.class.getName());

    @Override
    public XSSFWorkbook getXssfWorkbookFromFile(File file) throws InvalidSpreadSheetFormatException, IOException, InvalidFormatException {
        SpreadSheetFormat spreadSheetFormat;
        XSSFWorkbook workbook;
        try {
            spreadSheetFormat = ExcelUtilCore.getInstance().findXmlSpreadSheetFormat(file);

        } catch (Exception e){e.printStackTrace(); return null; }

        if (!spreadSheetFormat.equals(SpreadSheetFormat.XSSF)){
            throw  new InvalidSpreadSheetFormatException("Given file format is not XSSF format!");
        }
        workbook = new XSSFWorkbook(file);
        return  workbook;

    }

    @Override
    public List<XSSFSheet> getXssfSheetsFromFile(File file) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        List<XSSFSheet> sheets = new ArrayList<>();
        XSSFWorkbook workbook = getXssfWorkbookFromFile(file);
        int counterOfSheets = workbook.getNumberOfSheets();

        if (counterOfSheets<=0){ return  null; }

        for (int i =0; i<counterOfSheets;i++){
            sheets.add(workbook.getSheetAt(i));
        }
        return sheets;
    }

    @Override
    public XSSFSheet getXssfSheetFromFileByIndex(File file, int index) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFWorkbook workbook = getXssfWorkbookFromFile(file);

        if (index> workbook.getNumberOfSheets()){return null;}
        return  workbook.getSheetAt(index);
    }

    @Override
    public XSSFSheet getXssfSheetFromFileByName(File file, String sheet) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        XSSFWorkbook workbook =getXssfWorkbookFromFile(file);
        return  workbook.getSheet(sheet);
    }

    @Override
    public List<XSSFRow> getXssfRowsBySheet(XSSFSheet sheet) {
        List<XSSFRow> rows =new ArrayList<>();

        Iterator<Row> rowIterator =sheet.rowIterator();
        while (rowIterator.hasNext()){
            rows.add((XSSFRow)rowIterator.next());
        }
        return rows;
    }

    @Override
    public XSSFRow getXssfRowBySheetAndIndex(XSSFSheet sheet, int index) {
        if (sheet==null)
            return null;
        return sheet.getRow(index);
    }

    @Override
    public List<XSSFCell> getXssfCellsByRow(XSSFRow row) {
        if (row==null) return null;
        if (row.getPhysicalNumberOfCells()==0) return null;

        List<XSSFCell> cells =new ArrayList<>();

        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()){
            cells.add((XSSFCell) cellIterator.next());
        }
        return cells;

    }

    @Override
    public XSSFCell getXssfCellByRowAndIndex(XSSFRow row, int index) {
        if (row==null)return null;
        if (index> row.getPhysicalNumberOfCells()) return null;

        return  row.getCell(index);
    }



    @Override
    public List<Map<String, String>> parseXssfSheetToMapList(XSSFSheet sheet, boolean isFirstRowCellHeader,int startPos) {
        return null;
    }


    @Override
    public List<String> findCellHeaderNamesFromFirstRowXssfRowByRow(XSSFRow row,int startPos) {
        if (row==null)
        {
            logger.warning("Row is empty return null");
            return null;
        }
        XSSFRow targetRow;
        int cellCountOfTargetRow;
        List<String> headerList=new ArrayList<>();

        targetRow = row.getSheet().getRow(startPos);

        if (targetRow==null){
            logger.warning(String.format("row key : %s is empty row",startPos));
            return  null;
        }

        cellCountOfTargetRow=targetRow.getLastCellNum();
        logger.info(String.format("Searching Xssf cell header by row. Start row ->%s",startPos));
        for (int i =0;i<cellCountOfTargetRow;i++){

            headerList.add(targetRow.getCell(i).getStringCellValue());
        }

        logger.info(String.format("Document headers -> %s", String.join(",", headerList)));

        return  headerList;
    }

    @Override
    public List<String> findCellHeaderNamesByFirstXssfRowDefault(XSSFRow row) {
        return null;
    }

    @Override
    public Map<String, String> parseXssfRowToMap(XSSFRow row, boolean isFirstRowCellHeader,int startPos) {
        if (row==null)return null;

        Map<String,String> map = new HashMap<>();
        List<XSSFCell> cells=  getXssfCellsByRow(row);
        List<String> headerList=null ;
        int firstRowCellNumber=0;

        if(isFirstRowCellHeader){
            headerList = findCellHeaderNamesFromFirstRowXssfRowByRow(row, startPos);
        }else{
            firstRowCellNumber = row.getSheet().getRow(0).getLastCellNum();


        }

        return  null;
    }


}
