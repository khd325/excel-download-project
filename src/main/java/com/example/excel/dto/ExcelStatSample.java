package com.example.excel.dto;


import com.example.excel.ExcelBody;
import com.example.excel.ExcelHeader;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ExcelStatSample implements ExcelStatInterface{

    @ExcelHeader(headerName = "NO", colIndex = 0)
    @ExcelBody(rowIndex = 1, colIndex = 0)
    private Integer row;

    @ExcelHeader(headerName = "total", colIndex = 1)
    @ExcelBody(rowIndex = 1, colIndex = 1)
    private String total;

}
