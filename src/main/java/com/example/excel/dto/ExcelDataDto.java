package com.example.excel.dto;


import com.example.excel.ExcelBody;
import com.example.excel.ExcelHeader;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;

import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Getter
@Setter
@JsonIgnoreProperties(ignoreUnknown = true)
@AllArgsConstructor
@NoArgsConstructor
public class ExcelDataDto implements ExcelDtoInterface {


    @ExcelHeader(headerName = "이름",  colIndex = 0)
    @ExcelBody(rowIndex = 1, colIndex = 0)
    private String name;

    @ExcelHeader(headerName = "나이", colIndex = 1)
    @ExcelBody(rowIndex = 1, colIndex = 1)
    private int age;

    @ExcelHeader(headerName = "번호", colIndex = 2)
    @ExcelBody(rowIndex = 1, colIndex = 2)
    private String no;

    public static List<ExcelDataDto> makeDummyData() {


        return IntStream.rangeClosed(0, 10)
                .mapToObj(j -> {
                    ExcelDataDto testDto = new ExcelDataDto("name " + j, j, StringUtils.repeat(String.valueOf(j), 8));
                    return testDto;
                }).collect(Collectors.toList());

    }
}
