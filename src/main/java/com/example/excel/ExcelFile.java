package com.example.excel;

import com.example.excel.dto.ExcelInterface;
import com.example.excel.util.ExcelUtil;
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

    public void downloadExcel(HttpServletResponse response, String fileName, LinkedHashMap<String, ExcelInterface> dataMap, HashMap<String, Class> classInfoMap) throws ClassNotFoundException, IOException, IllegalAccessException {

        for (String key : dataMap.keySet()) {
            ExcelInterface data = dataMap.get(key);
            ExcelUtil excelUtil = new ExcelUtil();
//            if(data instanceof List) {
            if(data.getClass() == ArrayList.class) {
                List<ExcelInterface> dataList = (List<ExcelInterface>) data;
                Class<ExcelInterface> dtoClass = classInfoMap.get(key);

                excelUtil.export(workbook, dtoClass, dataList);

//                try {
//                    Method exportMethod = ExcelUtil.class.getMethod("export", SXSSFWorkbook.class, Class.class, List.class);
//                    exportMethod.invoke(new ExcelUtil(), workbook, dtoClass, dataList);
//                } catch (Exception e) {
//                    e.printStackTrace();
//                }
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
