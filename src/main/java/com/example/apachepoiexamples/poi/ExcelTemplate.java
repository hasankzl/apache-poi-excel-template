package com.example.apachepoiexamples.poi;


import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

@Data
public class ExcelTemplate {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private String filePath;
    private FileInputStream inputStream;
    private CellAddress address;
    private Cell  cellForUpdate;


    public ExcelTemplate(String filePath) throws IOException {

        this.filePath = filePath;
        ClassLoader classLoader = getClass().getClassLoader();
        inputStream = new FileInputStream(new File(Objects.requireNonNull(classLoader.getResource("static/"+filePath)).getFile()));
        workbook = new XSSFWorkbook(inputStream);
    }

    public void setSheet(Integer number){
        this.sheet = this.workbook.getSheetAt(number);
    }


    /**
     * Set value to a specific adress (like: A1,B3,G8)
     * If the cell at given address is empty that creates the cell in that address
     * @param cellAddress addressOf the cell
     * @param cellValue data to be written to the address
     */
    public void setValue(String cellAddress,String cellValue){

        address = new CellAddress(cellAddress);
        cellForUpdate = sheet.getRow(address.getRow()).getCell(address.getColumn());

        // if cell is null than create the cell
        if(cellForUpdate == null) {
            sheet.getRow(address.getRow()).createCell(address.getColumn());
            cellForUpdate = sheet.getRow(address.getRow()).getCell(address.getColumn());
        }
        cellForUpdate.setCellValue(cellValue);
    }

    private void closeInputStream() throws IOException {
        inputStream.close();
    }


    public ByteArrayOutputStream getOutputStream() throws IOException{
        this.closeInputStream();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        return outputStream;
    }
    /**
     * Automatic sizing of columns on the current sheet.
     */
    public void autoSizeSheetColumns(){
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i=0;i< numberOfSheets;i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            if(sheet.getPhysicalNumberOfRows()>0){
                Row row = sheet.getRow(sheet.getFirstRowNum());
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                }
            }
        }
    }

    /**
     * Shifts down the rows between startRow and the last row in excel by rowCount-1
     * After that, it copies the lines between copyStartRow and copyEndRow starting with startRow Copies up to rowCount
     * @param copyStartRow the index of the first row to copy
     * @param copyEndRow the index of the last row to copy
     * @param startRow the index of the starting number of copying
     * @param rowCount is the number of rows to copy
     */
    public void shiftAndCopyRows(Integer copyStartRow,Integer copyEndRow,Integer startRow,Integer rowCount){
        CellCopyPolicy cellCopyPolicy= new CellCopyPolicy();
        // shifting rows
        if(rowCount <=1){
            return;
        }
        else {
            sheet.shiftRows(startRow,sheet.getLastRowNum(),rowCount -1,true,true);
        }
        // after shifting copy row
        for (int i=0;i<rowCount -1;i++){
            sheet.copyRows(copyStartRow,copyEndRow,startRow+i,cellCopyPolicy);
        }
    }
    /**
     * Shifts down the rows between startRow and the last row in excel by rowCount-1
     * After that, it copies the lines between copyStartRow and copyEndRow starting with startRow Copies up to rowCount
     * @param dataList data to be written to cells
     * @param cellValues array of columns to write data to (like: "A","B","F")
     * @param startRow the index of the starting number of copying
     */
    public void fillRows(Integer startRow, String[] cellValues, List<Object[]> dataList){
        for(Object[] data: dataList){
            for(int i=0; i< data.length;i++){
                setValue(cellValues[i]+startRow,(String) data[i]);
            }
        }
    }


    /**
     * Get all merged regions between 2 columns
     * @param sheet Where to import merged regions
     * @param startCol the index of the start column
     * @param endCol the index of the end column
     */
    private List<CellRangeAddress> getMergedRegionsBetweenTwoColumns(XSSFSheet sheet, int startCol, int endCol) {
        List<CellRangeAddress> cellRangeAddressList = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() >= startCol && range.getLastColumn() <= endCol)
                cellRangeAddressList.add(range);
        }
        return cellRangeAddressList;
    }
    /**
     * Get all cells between 2 columns
     * @param sheet Where to import cells
     * @param startCol the index of the start column
     * @param endCol the index of the end column
     */
    private List<Cell> getCellsBetweenTwoColumns(XSSFSheet sheet, Integer startCol, Integer endCol) {
        List<Cell> cellList = new ArrayList<>();
        for (Row r : sheet) {
            for (int i = startCol; i <= endCol; i++) {
                Cell cell = r.getCell(i);
                if (cell != null)
                    cellList.add(cell);
                    // if cell equeals null than create the cell
                else{
                    r.createCell(i);
                    cellList.add(r.getCell(i));
                }
            }
        }
        return cellList;
    }
    /**
     * shift columns and insert copied columns
     * @param copyStartCol the starting index of the column to be copied
     * @param copyEndCol the ending index of the column to be copied
     * @param shiftStartCol the column index from which to start shifting
     * @param copyCount number of copy count
     */
    public void shiftAndCopyColumns(Integer copyStartCol,Integer copyEndCol,Integer shiftStartCol,Integer copyCount){
        Integer copyColumnCount = copyEndCol-copyStartCol+1;
        int shiftCount = copyCount *copyColumnCount;
        List<Cell> mainCellList =getCellsBetweenTwoColumns(sheet,copyStartCol,copyEndCol);
        List<CellRangeAddress> cellRangeAddressList = getMergedRegionsBetweenTwoColumns(sheet,copyStartCol,copyEndCol);
        sheet.shiftColumns(shiftStartCol,getSheetLastColNumber(),shiftCount);
        for (int i = 0; i < copyCount; i++) {
            for (CellRangeAddress range : cellRangeAddressList) {
                range.setFirstColumn(range.getFirstColumn() +copyColumnCount+ copyColumnCount * i);
                range.setLastColumn(range.getLastColumn() +copyColumnCount+ copyColumnCount * i);
                sheet.addMergedRegion(range);
            }
            for (int k = 0; k < mainCellList.size(); k++) {
                Cell oldCell = mainCellList.get(k);
                CellAddress cellAddress = oldCell.getAddress();
                Cell newCell = sheet.getRow(cellAddress.getRow()).createCell(cellAddress.getColumn()-copyStartCol+shiftStartCol+copyColumnCount * i);
                newCell.setCellStyle(oldCell.getCellStyle());
                newCell.setCellValue(oldCell.getStringCellValue());
            }
        }
    }
 /*
 * get max column number in current sheet
 * */
    private Integer getSheetLastColNumber(){
        Iterator rowIter = sheet.rowIterator();
        int col,maxcol =0;
        while (rowIter.hasNext()) {
            XSSFRow myRow = (XSSFRow) rowIter.next();
            //Cell iterator for iterating from cell to next cell of a row
            col = myRow.getLastCellNum();
            if(col >maxcol){
                maxcol = col;
            }
        }

        return maxcol;
    }
}
