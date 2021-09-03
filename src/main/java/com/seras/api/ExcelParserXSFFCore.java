package com.seras.api;

import com.seras.constants.LogConstants;
import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.interfaces.ExcelParserXSSF;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
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
    private ExcelParserXSFFCore(){}
    public static ExcelParserXSFFCore getInstance(){

        if (excelParserXSFFCore==null)
            excelParserXSFFCore=new ExcelParserXSFFCore();

        rowPointer =0;
        isFirstRowCellHeader=false;
        sheetPointer=0;
        return excelParserXSFFCore;
    }
    /* end singleton */

    private static int rowPointer =0;
    private static  boolean isFirstRowCellHeader=false;
    private static int sheetPointer =0;
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
    public List<Map<String, String>> parseXssfSheetToMapList(XSSFSheet sheet) {
        if (sheet==null)
            return null;

        List<Map<String,String>> rowListAsMap=new ArrayList<>();


        List<XSSFRow> xssfRowList = getXssfRowsBySheet(sheet);

        if (Optional.ofNullable(xssfRowList).map(List::size).orElse(0).equals(0)) {
            logger.warning(LogConstants.getInstance().msgRowListEmpty);
            return null;
        }
        logger.info("Start parsing rows to map...");
        xssfRowList.forEach(s->{
            Map<String,String> rowMap = parseXssfRowToMap(s);
            rowListAsMap.add(rowMap);
        });



        return rowListAsMap;
    }


    @Override
    public List<Map<String, String>> parseXssfWorkBookToMapList(XSSFWorkbook workbook) {
        XSSFSheet sheet =workbook.getSheetAt(sheetPointer);
        if(sheet==null){
            logger.warning(LogConstants.getInstance().msgSheetEmty);
            return null;
        }
        List<Map<String,String>> mapList = parseXssfSheetToMapList(sheet);
        return mapList;
    }

    @Override
    public List<Map<String, String>> parseExcelFileToMapList(File file) {
        return null;
    }

    @Override
    public List<String> findCellHeaderNamesFromFirstRowXssfRowByRow(XSSFRow row) {
        if (row==null)
        {
            logger.warning("Row is empty return null");
            return null;
        }
        XSSFRow targetRow;
        int cellCountOfTargetRow;
        List<String> headerList=new ArrayList<>();

        targetRow = row.getSheet().getRow(rowPointer);

        if (targetRow==null){
            logger.warning(String.format("row key : %s is empty row", rowPointer));
            return  null;
        }

        cellCountOfTargetRow=targetRow.getLastCellNum();
        logger.info(String.format("Searching Xssf cell header by row. Start row ->%s", rowPointer));
        for (int i =0;i<cellCountOfTargetRow;i++){
            String cellRawValue =targetRow.getCell(i).toString();
            headerList.add(cellRawValue);
        }

        logger.info(String.format("Document headers -> %s", String.join(",", headerList)));

        return  headerList;
    }

    @Override
    public List<String> findCellHeaderNamesByFirstXssfRowDefault(XSSFRow row) {
        if (row==null)
            return null;

        List<String> headerList =new ArrayList<>();
        int  firstRowCellNumber ;


        firstRowCellNumber=row.getSheet().getRow(rowPointer).getLastCellNum();

        for (int i =0;i<firstRowCellNumber;i++){
            headerList.add(CellReference.convertNumToColString(i));
        }

        return headerList;
    }

    @Override
    public Map<String, String> parseXssfRowToMap(XSSFRow row) {
        if (row==null)
        {
            logger.warning(LogConstants.getInstance().msgRowEmpty);
            return null;
        }

        Map<String,String> map = new HashMap<>();
        List<String> headerList ;


        if(isFirstRowCellHeader){
            headerList = findCellHeaderNamesFromFirstRowXssfRowByRow(row);
        }else{
            headerList =findCellHeaderNamesByFirstXssfRowDefault(row);
        }

        if(headerList==null)
        {
            logger.warning(LogConstants.getInstance().msgHeaderListEmpty);
            return null;
        }

        logger.info(String.format("Header List ->%s",String.join(",", headerList)));


        for (int i =0;i<headerList.size();i++){
            XSSFCell cell= row.getCell(i);
            String cellValueAsString = cell ==null ? "":cell.toString();
            map.put(headerList.get(i),cellValueAsString);
        }

        List<String> mapLogList =new ArrayList<>();
        map.forEach((s,v)-> mapLogList.add(String.format("\n  Cell Header : %s,Value : %s",s,v)));
        logger.info(String.format("Row Parsed To Map. %s",String.join(",",mapLogList)));
        return  map;
    }

    @Override
    public List<Map<String, String>> parseXssfFileToMapList(File file) throws IOException, InvalidFormatException, InvalidSpreadSheetFormatException {
        if(file==null){
            logger.warning(LogConstants.getInstance().msgFileEmpty);
            return null;
        }
        XSSFWorkbook workbook = getXssfWorkbookFromFile(file);
        return parseXssfWorkBookToMapList(workbook);
    }

    public ExcelParserXSFFCore setRowPointer(int startPointer){
        ExcelParserXSFFCore.rowPointer =startPointer;
        return this;
    }
    public  ExcelParserXSFFCore setIsFirstRowCellHeader(boolean isFirstRowCellHeader){
        ExcelParserXSFFCore.isFirstRowCellHeader = isFirstRowCellHeader;
        return this;
    }
    public ExcelParserXSFFCore setSheetPointer(int sheetPointer){
        ExcelParserXSFFCore.sheetPointer=sheetPointer;
        return this;
    }

}
