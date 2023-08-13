package com.example.excel.util;

import com.example.excel.ExcelBody;
import com.example.excel.ExcelHeader;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.*;
import java.lang.reflect.Field;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelUtil {


    public <T> void export(SXSSFWorkbook workbook, Class<T> excelClass, List<T> data) throws IllegalAccessException {
        SXSSFSheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<ExcelHeader>> headerMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelHeader.class))
                .map(field -> field.getDeclaredAnnotation(ExcelHeader.class))
                .sorted(Comparator.comparing(ExcelHeader::colIndex))
                .collect(Collectors.groupingBy(ExcelHeader::rowIndex));

        int index = getLastRow(sheet);


        for (Integer key : headerMap.keySet()) {
            SXSSFRow row = sheet.createRow(index++);
            for (ExcelHeader excelHeader : headerMap.get(key)) {
                SXSSFCell cell = row.createCell(excelHeader.colIndex());
                XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
                cell.setCellValue(excelHeader.headerName());

                cellStyle.setAlignment(excelHeader.headerStyle().horizontalAlignment());
                cellStyle.setVerticalAlignment(excelHeader.headerStyle().verticalAlignment());

                if (isHex(excelHeader.headerStyle().background().value())) {
                    cellStyle.setFillForegroundColor(new XSSFColor(Color.decode(excelHeader.headerStyle().background().value()), new DefaultIndexedColorMap()));
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }

                XSSFFont font = workbook.getXSSFWorkbook().createFont();
                font.setFontHeightInPoints((short) excelHeader.headerStyle().fontSize());
                font.setBold(true);
                cellStyle.setFont(font);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);

                cell.setCellStyle(cellStyle);
                if (excelHeader.colSpan() > 0 || excelHeader.rowSpan() > 0) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelHeader.rowSpan(), cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelHeader.colSpan());
                    sheet.addMergedRegion(cellAddresses);
                }
            }
        }

        Map<Integer, List<Field>> fieldMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class))
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                }).sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.groupingBy(field -> field.getDeclaredAnnotation(ExcelBody.class).rowIndex()));

        for (T t : data) {
            for (Integer key : fieldMap.keySet()) {
                SXSSFRow row = sheet.createRow(index++);
                for (Field field : fieldMap.get(key)) {
                    ExcelBody excelBody = field.getDeclaredAnnotation(ExcelBody.class);
                    Object o = field.get(t);
                    SXSSFCell cell = row.createCell(excelBody.colIndex());
                    XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();

                    cellStyle.setAlignment(excelBody.bodyStyle().horizontalAlignment());
                    cellStyle.setVerticalAlignment(excelBody.bodyStyle().verticalAlignment());

                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setBorderTop(BorderStyle.THIN);

                    cell.setCellValue(String.valueOf(o));

                    cell.setCellStyle(cellStyle);

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
        List<Field> groupField = Arrays.stream(excelClass.getDeclaredFields())
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
            for(int i = 0; i < groupMap.get(field).size(); i++) {
                SXSSFRow row = sheet.createRow(index++);
                SXSSFCell cell = row.getCell(field.getDeclaredAnnotation(ExcelBody.class).colIndex());
                if(!(dataRowIndex == groupMap.get(field).get(i))) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(dataRowIndex, groupMap.get(field).get(i), cell.getColumnIndex(), cell.getColumnIndex());
                    sheet.addMergedRegion(cellAddresses);
                }
                dataRowIndex = groupMap.get(field).get(i) + 1;
            }
        }

        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for(CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        }
    }


    private static boolean isHex(String hexCode) {
        if (StringUtils.startsWith(hexCode, "#")) {
            for (Character c : hexCode.substring(1).toCharArray()) {
                switch (c) {
                    case '0':
                    case '1':
                    case '2':
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    case '8':
                    case '9':
                    case 'a':
                    case 'b':
                    case 'c':
                    case 'd':
                    case 'e':
                    case 'f':
                    case 'A':
                    case 'B':
                    case 'C':
                    case 'D':
                    case 'E':
                    case 'F':
                        break;
                    default:
                        return false;
                }
            }
        } else {
            return false;
        }
        return true;
    }

    private int getLastRow(SXSSFSheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        return lastRowNum == -1 ? 0 : lastRowNum + 2;
    }
}
