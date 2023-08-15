package com.example.excel;

import com.example.excel.dto.ExcelDto;
import com.example.excel.util.ExcelUtil;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

public class ExcelFile {

    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;

    public ExcelFile() {
        this.workbook = new SXSSFWorkbook(1000);
        this.sheet = this.workbook.createSheet();
    }

    public void downloadExcel(HttpServletResponse response, String fileName, LinkedHashMap<String, Object> dataMap, HashMap<String, Class> classInfoMap) throws ClassNotFoundException, IOException {

        for (String key : dataMap.keySet()) {
            Object data = dataMap.get(key);

//            if(data instanceof List) {
            if(data.getClass() == ArrayList.class) {
                List<ExcelDto> dataList = (List<ExcelDto>) data;
                Class<ExcelDto> dtoClass = classInfoMap.get(key);

                try {
                    Method exportMethod = ExcelUtil.class.getMethod("export", SXSSFWorkbook.class, Class.class, List.class);
                    exportMethod.invoke(new ExcelUtil(), workbook, dtoClass, dataList);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }

        response.setHeader("Content-Disposition", "attachment; fileName=\"" + fileName + ".xlsx" + "\"");
        OutputStream os = response.getOutputStream();
        workbook.write(os);
        workbook.close();
        workbook.dispose();
        os.close();
    }
}
