package com.example.apachepoiexamples.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

public interface ExcelService {

    ByteArrayOutputStream getExcel() throws IOException;

    ByteArrayOutputStream shiftAndCopyRow() throws IOException;

    ByteArrayOutputStream shiftAndCopyRowFillData() throws IOException;

    ByteArrayOutputStream shiftAndCopyColumn() throws IOException;
}
