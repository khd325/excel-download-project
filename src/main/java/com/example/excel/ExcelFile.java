package com.example.excel;

import com.example.excel.util.ExcelUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;

public class ExcelFile {

    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;

    public ExcelFile() {
        this.workbook = new SXSSFWorkbook(1000);
        this.sheet = this.workbook.createSheet();
    }

    public void downloadExcel(HttpServletResponse response, String fileName, HashMap<String, Object> dataMap) throws ClassNotFoundException, IOException {

        for (String key : dataMap.keySet()) {
            Object data = dataMap.get(key);

            if(data instanceof List) {
                List<?> dataList = (List<?>) data;
                Class<?> dtoClass = dataList.get(0).getClass();

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
        os.close();
        workbook.dispose();
    }
}
