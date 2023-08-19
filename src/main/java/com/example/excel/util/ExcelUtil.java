package com.example.excel.util;

import com.example.excel.ExcelBody;
import com.example.excel.ExcelHeader;
import com.example.excel.HeaderStyle;
import com.example.excel.dto.ExcelDtoInterface;
import com.example.excel.dto.ExcelInterface;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;
import java.awt.Color;
import java.lang.reflect.Field;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelUtil {

    int[] columnLength;


    public void export(SXSSFWorkbook workbook, Class<ExcelInterface> excelClass, List<ExcelInterface> data) throws IllegalAccessException {
        SXSSFSheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<ExcelHeader>> headerMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelHeader.class))
                .map(field -> field.getDeclaredAnnotation(ExcelHeader.class))
                .sorted(Comparator.comparing(ExcelHeader::colIndex))
                .collect(Collectors.groupingBy(ExcelHeader::rowIndex));

        Field[] fields = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelHeader.class))
                .sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelHeader.class).colIndex()))
                .toArray(Field[]::new);

        Map<Integer, List<Field>> fieldMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class))
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                }).sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.groupingBy(field -> field.getDeclaredAnnotation(ExcelBody.class).rowIndex()));

        int index = getLastRow(sheet);

        if(columnLength == null) {
            columnLength = new int[fields.length];
            Arrays.fill(columnLength, 1);
        } else {
            if (columnLength.length < fields.length) {
                int[] temp = new int[fields.length];
                Arrays.fill(temp, 1);
                System.arraycopy(columnLength, 0, temp, 0, columnLength.length);
                columnLength = temp;
            }
        }

        XSSFCellStyle headerCellStyle = ExcelUtil.createHeaderCellStyle(workbook);
        XSSFCellStyle bodyCellStyle = ExcelUtil.createBodyCellStyle(workbook);

        for (Integer key : headerMap.keySet()) {
            SXSSFRow row = sheet.createRow(index++);
            for (ExcelHeader excelHeader : headerMap.get(key)) {
                SXSSFCell cell = row.createCell(excelHeader.colIndex());

                cell.setCellStyle(headerCellStyle);
                cell.setCellValue(excelHeader.headerName());

                if (excelHeader.colSpan() > 0 || excelHeader.rowSpan() > 0) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelHeader.rowSpan(), cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelHeader.colSpan());
                    sheet.addMergedRegion(cellAddresses);
                }
            }
        }

        for (ExcelInterface t : data) {
            for (Integer key : fieldMap.keySet()) {
                SXSSFRow row = sheet.createRow(index++);
                for (Field field : fieldMap.get(key)) {
                    ExcelBody excelBody = field.getDeclaredAnnotation(ExcelBody.class);
                    Object o = field.get(t);
                    SXSSFCell cell = row.createCell(excelBody.colIndex());


                    if(o == null) {
                        cell.setCellValue("");
                    } else if (o.getClass() == Integer.class) {
                        cell.setCellValue(Integer.parseInt(String.valueOf(o)));
                    } else {
                        cell.setCellValue(String.valueOf(o));
                    }

                    cell.setCellStyle(bodyCellStyle);

                    if (excelBody.colSpan() > 0 || excelBody.rowSpan() > 0) {
                        CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelBody.rowSpan(),
                                cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelBody.colSpan());
                        sheet.addMergedRegion(cellAddresses);
                    }

                    if ((excelBody.width() > 0 && excelBody.width() != 8) && sheet.getColumnWidth(excelBody.colIndex()) == 2048) {
                        sheet.setColumnWidth(excelBody.colIndex(), excelBody.width() * 256);
                    }
                }
            }
        }
        /*List<Field> groupField = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class) && field.getDeclaredAnnotation(ExcelBody.class).rowGroup())
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                }).sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.toList());

        Map<Field, List<Integer>> groupMap = new HashMap<>();
        for (Field field : groupField) {
            groupMap.put(field, new ArrayList<>());
            for (int i = 0; i < data.size(); i++) {
                Object o1 = field.get(data.get(i));

                for (int j = i + 1; j < data.size(); j++) {
                    Object o2 = field.get(data.get(j));

                    if (!o1.equals(o2)) {
                        groupMap.get(field).add((j) * headerMap.size() + headerMap.keySet().size() - 1);
                        i = j - 1;
                        break;
                    }
                }
            }
            groupMap.get(field).add(sheet.getLastRowNum());
        }

        for (Field field : groupMap.keySet()) {
            int dataRowIndex = headerMap.keySet().size();
            for (int i = 0; i < groupMap.get(field).size(); i++) {
                SXSSFRow row = sheet.createRow(index++);
                SXSSFCell cell = row.getCell(field.getDeclaredAnnotation(ExcelBody.class).colIndex());
                if (!(dataRowIndex == groupMap.get(field).get(i))) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(dataRowIndex, groupMap.get(field).get(i), cell.getColumnIndex(), cell.getColumnIndex());
                    sheet.addMergedRegion(cellAddresses);
                }
                dataRowIndex = groupMap.get(field).get(i) + 1;
            }
        }

        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        }*/
    }


    private int getLastRow(SXSSFSheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        return lastRowNum == -1 ? 0 : lastRowNum + 2;
    }

    public static XSSFCellStyle createHeaderCellStyle(SXSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)9);
        font.setFontName("맑은 고딕");


        XSSFCellStyle cellStyle = workbook.getXSSFWorkbook().createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFont(font);

        XSSFColor xssfColor = new XSSFColor(Color.decode("#F0F0F0"), new DefaultIndexedColorMap());
        cellStyle.setFillForegroundColor(xssfColor);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);


        return cellStyle;
    }

    public static XSSFCellStyle createBodyCellStyle(SXSSFWorkbook workbook) {

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("맑은 고딕");

        XSSFCellStyle cellStyle = workbook.getXSSFWorkbook().createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFont(font);

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);

        return cellStyle;
    }
}
