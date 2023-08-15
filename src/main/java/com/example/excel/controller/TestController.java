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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

@RestController
public class TestController {

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException, IllegalAccessException {

//        List<TestDto> testDtos = TestDto.makeDummyData();
//        testDtos.add(new TestDto(null,19));
//        List<TestDto.StatDto> statDtos = TestDto.StatDto.makeDummyData();

        List<TestDto> testDtos = new ArrayList<>();
        List<TestDto.StatDto> statDtos = new ArrayList<>();

        LinkedHashMap<String, Object> dataMap = new LinkedHashMap<>();
        HashMap<String, Class> classInfoMap = new HashMap<>();

        dataMap.put("stat", statDtos);
        classInfoMap.put("stat", TestDto.StatDto.class);
        dataMap.put("dataList", testDtos);
        classInfoMap.put("dataList", TestDto.class);
        ExcelFile excelFile = new ExcelFile();
        try {
            excelFile.downloadExcel(response, "test", dataMap, classInfoMap);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

}
