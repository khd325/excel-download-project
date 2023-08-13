package com.example.excel.controller;

import com.example.excel.ExcelFile;
import com.example.excel.dto.TestDto;
import com.example.excel.util.ExcelUtil;
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
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

@RestController
public class TestController {

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException, IllegalAccessException {

        List<TestDto> testDtos = TestDto.makeDummyData();
        testDtos.add(new TestDto(null,19));
        List<TestDto.StatDto> statDtos = TestDto.StatDto.makeDummyData();

        LinkedHashMap<String, Object> dataMap = new LinkedHashMap<>();
//        dataMap.put("stat", statDtos);
        dataMap.put("dataList", testDtos);

        ExcelFile excelFile = new ExcelFile();
        try {
            excelFile.downloadExcel(response, "test", dataMap);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

}
