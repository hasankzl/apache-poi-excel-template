package com.example.apachepoiexamples.resource;


import com.example.apachepoiexamples.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.util.IOUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;

    @GetMapping
    public void getExcelFile(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition","attachment;filename=poi.xlsx");
        ByteArrayInputStream stream = new ByteArrayInputStream(excelService.getExcel().toByteArray());
        IOUtils.copy(stream,response.getOutputStream());
    }


    @GetMapping("/shiftAndCopyRow")
    public void shiftAndCopyRow(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition","attachment;filename=poi.xlsx");

        ByteArrayInputStream stream = new ByteArrayInputStream(excelService.shiftAndCopyRow().toByteArray());
        IOUtils.copy(stream,response.getOutputStream());
    }

    @GetMapping("/shiftAndCopyRowFillData")
    public void shiftAndCopyRowFillData(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition","attachment;filename=poi.xlsx");

        ByteArrayInputStream stream = new ByteArrayInputStream(excelService.shiftAndCopyRowFillData().toByteArray());
        IOUtils.copy(stream,response.getOutputStream());
    }



    @GetMapping("/shiftAndCopyColumn")
    public void shiftAndCopyColumn(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition","attachment;filename=poi.xlsx");
        ByteArrayInputStream stream = new ByteArrayInputStream(excelService.shiftAndCopyColumn().toByteArray());
        IOUtils.copy(stream,response.getOutputStream());
    }
}
