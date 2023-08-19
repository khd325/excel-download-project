package com.example.excel.controller;

import com.example.excel.ExcelFile;
import com.example.excel.dto.ExcelDataDto;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

@RestController
public class TestController {

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException, IllegalAccessException {

        List<ExcelDataDto> testDtos = ExcelDataDto.makeDummyData();

        LinkedHashMap<String, Object> dataMap = new LinkedHashMap<>();
        HashMap<String, Class> classInfoMap = new HashMap<>();

        dataMap.put("dataList", testDtos);
        classInfoMap.put("dataList", ExcelDataDto.class);
        ExcelFile excelFile = new ExcelFile();
        try {
            excelFile.downloadExcel(response, "test", dataMap, classInfoMap);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

}
