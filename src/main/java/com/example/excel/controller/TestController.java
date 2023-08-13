package com.example.excel.controller;

import com.example.excel.dto.TestDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.IOException;
import java.util.List;

@RestController
public class TestController {

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new SXSSFWorkbook();

        CellStyle greyCellStyle = workbook.createCellStyle();
        applyCellStyle(greyCellStyle, new Color(231,234,236));

        CellStyle bodyCellStyle = workbook.createCellStyle();
        applyCellStyle(bodyCellStyle, new Color(255,255,255));

        Sheet sheet = workbook.createSheet();

        List<TestDto> testDtos = TestDto.makeDummyData();

        int rowIndex = 0;
        Row headerRow = sheet.createRow(rowIndex++);

        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellStyle(greyCellStyle);
        headerCell.setCellValue("이름");

        headerCell = headerRow.createCell(1);
        headerCell.setCellStyle(greyCellStyle);
        headerCell.setCellValue("나이");

        for (TestDto testDto : testDtos) {
            Row bodyRow = sheet.createRow(rowIndex++);

            Cell bodyCell = bodyRow.createCell(0);
            bodyCell.setCellStyle(bodyCellStyle);
            bodyCell.setCellValue(testDto.getName());

            bodyCell = bodyRow.createCell(1);
            bodyCell.setCellStyle(bodyCellStyle);
            bodyCell.setCellValue(testDto.getAge());
        }

        response.setContentType("application/vnd.ms-excel");

        workbook.write(response.getOutputStream());
        workbook.close();
    }


    private void applyCellStyle(CellStyle cellStyle, java.awt.Color color) {
        XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
        xssfCellStyle.setFillForegroundColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
    }
}
