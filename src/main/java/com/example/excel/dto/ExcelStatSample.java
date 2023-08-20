package com.example.excel.dto;


import com.example.excel.ExcelBody;
import com.example.excel.ExcelHeader;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ExcelStatSample implements ExcelStatInterface{

    @ExcelHeader(headerName = "NO", colIndex = 1)
    @ExcelBody(colIndex = 1, rowIndex = 1)
    private Integer row;

    @ExcelHeader(headerName = "summary", colIndex = 0, rowIndex = 1)
    private String summary;

    @ExcelHeader(headerName = "total", colIndex = 2)
    @ExcelBody(colIndex = 2, rowIndex = 1)
    private String total;


    public static void makeDummy(List<ExcelInterface> list) {

        list.add(new ExcelStatSample(1,null, "12341567"));
    }

}
