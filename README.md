#  apache-poi-excel-template

Includes useful functions for updating excel file with apache poi  like adding rows, adding cols etc.


# You can find answers for these questions in that repository.
 

 ### How to download excel file?  
  
 ### How to update a excel file with apach poi?  
  
 ### How to add rows in a excel file?  
  
 ### How to add rows with data in a excel file?  
  
 ### How to change a celle is values by cell address?  
  
 ### How to copy and paste columns in a excel file ? (merged cells included)

## How to work with ExcelTemplate class?
ExcelTemplate class take a paramether witch is the file name.

    ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
(ExcelTemplate tryies to find your excel file in resource you can look at constroctor function of ExcelTemplate)

### Do not forget to set sheet that you wanna work on !

After create ExcelTemplate you need to set the sheet that you wanna work on. By defaut it will be set to 0.

    excelTemplate.setSheet(0);

after that you can update your excel file with excelTemplate class.
## How to download excel file?  

After you create a ExcelTemplate you can get output stream 

> excelTemplate.getOutputStream();

In controller you can return the file with response

    @GetMapping  
    public void getExcelFile(HttpServletResponse response) throws IOException {  
    ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
    response.setContentType("application/octet-stream");  
    response.setHeader("Content-Disposition","attachment;filename=poi.xlsx");  
    ByteArrayInputStream stream = new ByteArrayInputStream( excelTemplate.getOutputStream().toByteArray());      
    IOUtils.copy(stream,response.getOutputStream());  
    }

You can check resource/ExcelController for more examples.


## How to add rows in a excel file?  

For adding row you can use shiftAndCopyRows() function. This function will shift the rows and copy a column in that shifted places. For example lets assume that you have a excel file like this
![addrowexample](https://i.imgur.com/IWDEmO8.png)
 and you wanna add new rows in that table. To do that you can simply use shiftAndCopyRows().
 

    public void shiftAndCopyRows(Integer copyStartRow,Integer copyEndRow,Integer startRow,Integer rowCount)
 * @param copyStartRow the index of the first row to copy  
* @param copyEndRow the index of the last row to copy  
* @param startRow the index of the starting number of copying  
* @param rowCount is the number of rows to copy

after use that our file will be like this

      excelTemplate.shiftAndCopyRows(1,2,2,4);

 ![addRow2](https://i.imgur.com/tamuIdC.png)



## How to add rows with data in a excel file?

You can do that with combining 2 functions shiftAndCopyRows() and fillRows()
First shift and copy rows like above example. After that use fillRow() function to fill that rows.

    public void fillRows(Integer startRow, String[] cellValues, List<Object[]> dataList)
* @param cellValues array of columns to write data to (like: "A","B","F")  
* @param startRow the index of the starting number of copying  
* @param dataList data to be written to cells

first define cellValues and create dataList 

    String[] cellValues = {"A","B","C","D","E"};
    List<Object[]> dataList = new ArrayList<>();  
    Object[] obj = new Object[]{  
        "hasan",  
    "kuzulu",  
     "00000000",  
    "Istanbul",  
    "mhasan.kzl@gmail.com"  
    };  
  
    Object[] obj2 = new Object[]{  
        "Julian ",  
     "alexander",  
     "00000000",  
     "USA",  
     "julian.alexander@example.com"  
    };  
    dataList.add(obj);  
    dataList.add(obj2);

Our data and cellValues is ready, now lets use this values in the function.

    excelTemplate.shiftAndCopyRows(1,2,2,dataList.size());  
    excelTemplate.fillRows(2,cellValues,dataList);

first shift and copy rows by my dataList count after that paste the data in that rows.
Finally our files look like this
![addingRowWithData](https://i.imgur.com/L1Vbxeo.png)

## How to change a celle is values by cell address?  

This is really simple just use setValue() function. First paramether is the cell Address (like : "A1","C3","D5") second is the value.

    excelTemplate("A4","HELLO WORLD")


#  How to copy and paste columns in a excel file ?

For copy and paste columns you can use shiftAndCopyColumns() function. 
Lets assume that you have a excel file like this.

![pasteCol1](https://i.imgur.com/e61IX3w.png)

Somehow you need to copy surname,tel,city columns 2 times before address column.
To do that you can simply use shiftAndCopyColumns() function like this.
* @param copyStartCol the starting index of the column to be copied  
* @param copyEndCol the ending index of the column to be copied  
* @param shiftStartCol the column index from which to start shifting  
* @param copyCount number of copy count

      public void shiftAndCopyColumns(Integer copyStartCol,Integer copyEndCol,Integer shiftStartCol,Integer    copyCount)
Now letst coppy columns between 1. and 3. column and paste in 3 times starting with 4. column

      excelTemplate.shiftAndCopyColumns(1,3,4 ,3);

After that our file will be updated to this
![addingCol2](https://i.imgur.com/E5aSOXf.png)
