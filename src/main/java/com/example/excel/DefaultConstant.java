package com.example.excel;

import lombok.Getter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

public enum DefaultConstant {


    GREY("#ECEFF3");

    private final String color;

    public String color() {
        return color;
    }


    DefaultConstant(String color) {
        this.color = color;
    }
}
