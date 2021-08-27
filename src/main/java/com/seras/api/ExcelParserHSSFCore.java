package com.seras.api;

import com.seras.constants.LogConstants;
import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.interfaces.ExcelParserHSSF;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.logging.Logger;

public class ExcelParserHSSFCore implements ExcelParserHSSF {

    /* begin singleton */
    private static ExcelParserHSSFCore excelParserHSSFCore;
    private  ExcelParserHSSFCore(){}
    public static ExcelParserHSSFCore getInstance(){
        if (excelParserHSSFCore==null)
            excelParserHSSFCore=new ExcelParserHSSFCore();

        rowPointer=0;
        isFirstRowCellHeader=false;
        sheetPointer=0;
        return excelParserHSSFCore;
    }
    /* end singleton */

    private static int rowPointer =0;
    private static  boolean isFirstRowCellHeader=false;
    private static int sheetPointer =0;

    Logger logger = Logger.getLogger(ExcelParserHSSFCore.class.getName());
    @Override
    public HSSFWorkbook getHssfWorkbookFromFile(File file) throws InvalidSpreadSheetFormatException, IOException {

        SpreadSheetFormat spreadSheetFormat;
        try {
            spreadSheetFormat = ExcelUtilCore.getInstance().findXmlSpreadSheetFormat(file);
        } catch (Exception e){e.printStackTrace(); return null; }

        if (!spreadSheetFormat.equals(SpreadSheetFormat.HSSF)){
            throw  new InvalidSpreadSheetFormatException("Given file format is not HSSF format!");
        }
        FileInputStream fileInputStream = new FileInputStream(file);
        return  new HSSFWorkbook(fileInputStream);
    }

    @Override
    public List<HSSFSheet> getHssfSheetsFromFile(File file) throws IOException, InvalidSpreadSheetFormatException {
        List<HSSFSheet> sheets= new ArrayList<>();
        HSSFWorkbook workbook = getHssfWorkbookFromFile(file);
        int counterOfSheets = workbook.getNumberOfSheets();

        if (counterOfSheets<=0){ return null; }

        for (int i =0; i<counterOfSheets;i++){
            sheets.add(workbook.getSheetAt(i));
        }

        return  sheets;
    }

    @Override
    public HSSFSheet getHssfSheetFromFileByIndex(File file, int index) throws IOException, InvalidSpreadSheetFormatException {
        HSSFWorkbook workbook = getHssfWorkbookFromFile(file);

        if(index> workbook.getNumberOfSheets()){return null;}
        return workbook.getSheetAt(index);

    }
    @Override
    public HSSFSheet getHssfSheetFromFileByName(File file, String sheet) throws IOException, InvalidSpreadSheetFormatException {
        HSSFWorkbook workbook =getHssfWorkbookFromFile(file);
        return workbook.getSheet(sheet);
    }

    @Override
    public List<HSSFRow> getHssfRowsBySheet(HSSFSheet sheet) {
        List<HSSFRow> rows =new ArrayList<>();

        Iterator<Row> rowIterator =sheet.rowIterator();
        while (rowIterator.hasNext()){
            rows.add((HSSFRow) rowIterator.next());
        }
        return  rows;
    }

    @Override
    public HSSFRow getHssfRowBySheetAndIndex(HSSFSheet sheet, int index) {
        if (sheet==null)
            return null;
        return  sheet.getRow(index);
    }

    @Override
    public List<HSSFCell> getHssfCellsByRow(HSSFRow row) {
        if (row==null)return null;

        List<HSSFCell> cells=new ArrayList<>();
        Iterator<Cell> cellIterator=row.cellIterator();

        while (cellIterator.hasNext()){
            cells.add((HSSFCell) cellIterator.next());
        }
        return cells;

    }

    @Override
    public HSSFCell getHssfCellByRowAndIndex(HSSFRow row, int index) {
        if (row==null)return  null;
        if (index>row.getPhysicalNumberOfCells())return null;

        return row.getCell(index);
    }

    @Override
    public List<Map<String, String>> parseHssfSheetToMapList(HSSFSheet sheet) {
        if (sheet==null)
            return null;

        List<Map<String,String>> rowListAsMap=new ArrayList<>();


        List<HSSFRow> hssfRowList = getHssfRowsBySheet(sheet);

        if (Optional.ofNullable(hssfRowList).map(List::size).orElse(0).equals(0)) {
            logger.warning(LogConstants.getInstance().msgRowListEmpty);
            return null;
        }
        logger.info("Start parsing rows to map...");
        hssfRowList.forEach(s->{
            Map<String,String> rowMap = parseHssfRowToMap(s);
            rowListAsMap.add(rowMap);
        });

        return rowListAsMap;
    }


    @Override
    public Map<String, String> parseHssfRowToMap(HSSFRow row) {
        if (row==null)
        {
            logger.warning(LogConstants.getInstance().msgRowEmpty);
            return null;
        }

        Map<String,String> map = new HashMap<>();
        List<String> headerList ;


        if(isFirstRowCellHeader){
            headerList = findCellHeaderNamesFromFirstRowHssfRowByRow(row);
        }else{
            headerList =findCellHeaderNamesByFirstHssfRowDefault(row);
        }

        if(headerList==null)
        {
            logger.warning(LogConstants.getInstance().msgHeaderListEmpty);
            return null;
        }

        logger.info(String.format("Header List ->%s",String.join(",", headerList)));


        for (int i =0;i<headerList.size();i++){
            HSSFCell cell= row.getCell(i);
            String cellValueAsString = cell ==null ? "":cell.toString();
            map.put(headerList.get(i),cellValueAsString);
        }

        List<String> mapLogList =new ArrayList<>();
        map.forEach((s,v)-> mapLogList.add(String.format("\n  Cell Header : %s,Value : %s",s,v)));
        logger.info(String.format("Row Parsed To Map. %s",String.join(",",mapLogList)));
        return  map;
    }

    @Override
    public List<Map<String, String>> parseHssfWorkBookToMapList(HSSFWorkbook workbook) {
        if(workbook ==null){
            logger.warning(LogConstants.getInstance().msgWorkbookEmpty);
            return null;
        }
        HSSFSheet sheet = workbook.getSheetAt(sheetPointer);
        List<Map<String,String>> maps=parseHssfSheetToMapList(sheet);
        return maps;


    }

    @Override
    public List<Map<String, String>> parseHssfFileToMapList(File file) throws IOException, InvalidSpreadSheetFormatException {
        if(file ==null){
            logger.warning(LogConstants.getInstance().msgFileEmpty);
            return null;
        }
        HSSFWorkbook workbook = getHssfWorkbookFromFile(file);
        return parseHssfWorkBookToMapList(workbook);
    }

    @Override
    public List<String>findCellHeaderNamesFromFirstRowHssfRowByRow(HSSFRow row) {
        if (row==null) return null;
        HSSFRow targetRow;
        int cellCountOfTargetRow;
        List<String> headerList=new ArrayList<>();

        targetRow=row.getSheet().getRow(rowPointer);

        if (targetRow==null) return null;

        cellCountOfTargetRow =targetRow.getLastCellNum();
        logger.info(String.format("Searching Hssf cell header by row. Start row ->%s",rowPointer));

        for (int i =0;i<cellCountOfTargetRow;i++){

            headerList.add(targetRow.getCell(i).getStringCellValue());
        }
        logger.info(String.format("Document headers -> %s", String.join(",", headerList)));

        return  headerList;

    }


    @Override
    public List<String> findCellHeaderNamesByFirstHssfRowDefault(HSSFRow row) {

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

    public ExcelParserHSSFCore setRowPointer(int startPointer){
        ExcelParserHSSFCore.rowPointer =startPointer;
        return this;
    }
    public  ExcelParserHSSFCore setIsFirstRowCellHeader(boolean isFirstRowCellHeader){
        ExcelParserHSSFCore.isFirstRowCellHeader = isFirstRowCellHeader;
        return this;
    }
    public ExcelParserHSSFCore setSheetPointer(int sheetPointer){
        ExcelParserHSSFCore.sheetPointer=sheetPointer;
        return this;
    }
}
