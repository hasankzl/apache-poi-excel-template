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
        ExcelTemplate excelTemplate = new ExcelTemplate("static/test.xlsx");
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyRow() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("static/testAddingRow.xlsx");
        excelTemplate.setSheet(0);
        excelTemplate.shiftAndCopyRows(1,2,2,4);
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyRowFillData() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("static/testAddingRow.xlsx");
        excelTemplate.setSheet(0);
        String[] cellValues = {"A","B","C","D","E","F"};
        List<Object[]> dataList = new ArrayList<>();
        Object[] obj = new Object[]{
                "hasan",
                "kuzulu",
                112321321,
                12.3,
                "Istanbul",
                "mhasan.kzl@gmail.com"
        };

        Object[] obj2 = new Object[]{
                "Julian ",
                "alexander",
                123123123,
                17.3,
                "USA",
                "julian.alexander@example.com"
        };
        dataList.add(obj);
        dataList.add(obj2);
        excelTemplate.shiftAndCopyRows(1,2,2,dataList.size());
        excelTemplate.fillRows(2,cellValues,dataList);
        return excelTemplate.getOutputStream();
    }

    @Override
    public ByteArrayOutputStream shiftAndCopyColumn() throws IOException {
        ExcelTemplate excelTemplate = new ExcelTemplate("static/test.xlsx");
        excelTemplate.setSheet(0);

        excelTemplate.shiftAndCopyColumns(1,3,4 ,3);

        return excelTemplate.getOutputStream();
    }
}
