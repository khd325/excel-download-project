package com.example.excel.dto;


import com.example.excel.*;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class TestDto extends ExcelDto {


    @ExcelHeader(headerName = "이름",  colIndex = 0)
    @ExcelBody(rowIndex = 1, colIndex = 0, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private String name;

    @ExcelHeader(headerName = "나이", colIndex = 1, colSpan = 1)
    @ExcelBody(rowIndex = 1, colIndex = 1, colSpan = 1, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private int age;

    @ExcelHeader(headerName = "번호", colIndex = 3)
    @ExcelBody(rowIndex = 1, colIndex = 3)
    private String no;




    @AllArgsConstructor
    @NoArgsConstructor
    @Getter
    @Setter
    public static class StatDto {

        @ExcelHeader(headerName = "테스트1", colIndex = 0)
        @ExcelBody(rowIndex = 1, colIndex = 0, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
        private int test1;


        @ExcelHeader(headerName = "테스트2", colIndex = 1)
        @ExcelBody(rowIndex = 1, colIndex = 1, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
        private String test2;

        public static List<StatDto> makeDummyData() {

            ArrayList<StatDto> arr = new ArrayList<>();
            arr.add(new StatDto(11, "테스트2"));
            return arr;
        }
    }


    public static List<TestDto> makeDummyData() {


        return IntStream.rangeClosed(0, 10)
                .mapToObj(j -> {
                    TestDto testDto = new TestDto("name " + j, j, StringUtils.repeat(String.valueOf(j), 8));
                    return testDto;
                }).collect(Collectors.toList());

    }
}
