package com.example.excel.controller;

import com.example.excel.dto.TestDto;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@RestController
public class TestController {

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new SXSSFWorkbook();

        Sheet sheet = workbook.createSheet();

        List<TestDto> testDtos = TestDto.makeDummyData();

        int rowIndex = 0;
        Row headerRow = sheet.createRow(rowIndex++);

        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("이름");

        headerCell = headerRow.createCell(1);
        headerCell.setCellValue("나이");

        for (TestDto testDto : testDtos) {
            Row bodyRow = sheet.createRow(rowIndex++);

            Cell bodyCell = bodyRow.createCell(0);
            bodyCell.setCellValue(testDto.getName());

            bodyCell = bodyRow.createCell(1);
            bodyCell.setCellValue(testDto.getAge());
        }

        response.setContentType("application/vnd.ms-excel");

        workbook.write(response.getOutputStream());
        workbook.close();


    }
}
