package com.example.apachepoiexamples.service.impl;

import com.example.apachepoiexamples.poi.ExcelTemplate;
import com.example.apachepoiexamples.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {
    @Override
    public ByteArrayOutputStream getExcel() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyRow() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
        excelTemplate.setSheet(0);
        excelTemplate.shiftAndCopyRows(1,2,2,4);
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyRowFillData() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
        excelTemplate.setSheet(0);
        String[] fillRows = {"A","B","C","D"};
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
        excelTemplate.shiftAndCopyRows(1,2,2,dataList.size());
        excelTemplate.fillRows(2,fillRows,dataList);
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyColumn() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("test.xlsx");
        excelTemplate.setSheet(0);

        excelTemplate.shiftAndCopyColumns(1,3,4 ,3);

        return excelTemplate.getOutputStream();
    }
}
