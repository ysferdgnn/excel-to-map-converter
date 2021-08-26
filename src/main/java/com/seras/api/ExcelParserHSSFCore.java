package com.seras.api;

import com.seras.enums.SpreadSheetFormat;
import com.seras.exceptions.InvalidSpreadSheetFormatException;
import com.seras.interfaces.ExcelParserHSSF;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class ExcelParserHSSFCore implements ExcelParserHSSF {

    /* begin singleton */
    private static ExcelParserHSSFCore excelParserHSSFCore;
    private  ExcelParserHSSFCore(){}
    public static ExcelParserHSSFCore getInstance(){
        if (excelParserHSSFCore==null)
            excelParserHSSFCore=new ExcelParserHSSFCore();
        return excelParserHSSFCore;
    }
    /* end singleton */

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
    public List<Map<String, String>> parseHssfSheetToMapList(HSSFSheet sheet, boolean isFirstRowCellHeader, int startPos) {
        return null;
    }


    @Override
    public Map<String, String> parseHssfRowToMap(HSSFRow row, boolean isFirstRowCellHeader, int startPos) {
        return null;
    }

    @Override
    public List<String>findCellHeaderNamesFromFirstRowHssfRowByRow(HSSFRow row,int startPos) {
        if (row==null) return null;
        HSSFRow targetRow;
        int cellCountOfTargetRow;
        List<String> headerList=new ArrayList<>();

        targetRow=row.getSheet().getRow(startPos);

        if (targetRow==null) return null;

        cellCountOfTargetRow =targetRow.getLastCellNum();
        logger.info(String.format("Searching Hssf cell header by row. Start row ->%s",startPos));

        for (int i =0;i<cellCountOfTargetRow;i++){

            headerList.add(targetRow.getCell(i).getStringCellValue());
        }
        logger.info(String.format("Document headers -> %s", String.join(",", headerList)));

        return  headerList;

    }

    @Override
    public List<String> findCellHeaderNamesByFirstHssfRowDefault(HSSFRow row) {
        return null;
    }
}
