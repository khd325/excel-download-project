package com.example.excel.dto;


import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class TestDto {

    private String name;

    private int age;


    public static List<TestDto> makeDummyData() {


        return IntStream.rangeClosed(0, 10)
                .mapToObj(j -> {
                    TestDto testDto = new TestDto("name " + j, j);
                    return testDto;
                }).collect(Collectors.toList());

    }
}
