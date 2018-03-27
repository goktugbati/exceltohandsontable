package com.excel.demo.service;

import com.excel.demo.model.SheetModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelPOIHelper {

    private String readCellContent(Cell cell) {
        String content;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                content = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    content = cell.getDateCellValue() + "";
                } else {
                    content = cell.getNumericCellValue() + "";
                }
                break;
            case BOOLEAN:
                content = cell.getBooleanCellValue() + "";
                break;
            case FORMULA:
                content = cell.getCellFormula() + "";
                break;
            default:
                content = "";
        }
        return content;
    }


    public List<SheetModel> getSheetNames(String fileLocation) throws IOException {
        //Converts File to an InputStream
        FileInputStream fis = new FileInputStream(new File(fileLocation));
        //Creates a workbook and sets it to null
        XSSFWorkbook workbook = null;
        //Key holds Sheet number, Value holds Sheet name
        workbook = new XSSFWorkbook(fis);
        List<SheetModel> sheetModelList = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++){
            SheetModel sheetModel = new SheetModel();
            sheetModel.setSheetNo(i);
            sheetModel.setSheetName(workbook.getSheetName(i));
            sheetModelList.add(sheetModel);
        }

        return sheetModelList;
    }



    //Reads your excel file like a BOSS so you can display what's in it.
    public Map<Integer, List<String>> readMyExcelFileLikeaBOSS(String fileLocation, int sheetNumber) throws IOException {
        //Converts File to an InputStream
        FileInputStream fis = new FileInputStream(new File(fileLocation));
        //Creates a workbook and sets it to null
        XSSFWorkbook workbook = null;
        //Creates HashMap object to write data from the excel file
        Map<Integer, List<String>> data = new HashMap<>();
        workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();

            for (int i = firstRowNum; i < lastRowNum; i++) {
                XSSFRow row;
                row = sheet.getRow(i);

                data.put(i, new ArrayList<>());

                if (row != null) {
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        if (cell != null) {
                            String myCellData = readCellContent(cell);
                                data.get(i)
                                    .add(myCellData);
                        } else {
                            data.get(i).add("");
                        }
                    }
                }
            }

        return data;
    }


    public void updateExcel(int sheetNumber, String cellContent, int rowNumber, int cellNumber){

    }

}
