package com.example.excel.dto;


import com.example.excel.*;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class TestDto extends ExcelDto {


    @ExcelHeader(headerName = "이름",  colIndex = 0, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 0, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private String name;

    @ExcelHeader(headerName = "나이", colIndex = 1, colSpan = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 1, colSpan = 1, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private int age;




    @AllArgsConstructor
    @NoArgsConstructor
    @Getter
    @Setter
    public static class StatDto {

        @ExcelHeader(headerName = "테스트1", colIndex = 0, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
        @ExcelBody(rowIndex = 1, colIndex = 0, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
        private int test1;


        @ExcelHeader(headerName = "테스트2", colIndex = 1, headerStyle = @HeaderStyle(background = @Background("#ECEFF3")))
        @ExcelBody(rowIndex = 1, colIndex = 1, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
        private String test2;

        public static List<StatDto> makeDummyData() {

            return List.of(new StatDto(1000, "test2"));
        }
    }


    public static List<TestDto> makeDummyData() {


        return IntStream.rangeClosed(0, 10)
                .mapToObj(j -> {
                    TestDto testDto = new TestDto("name " + j, j);
                    return testDto;
                }).collect(Collectors.toList());

    }
}
