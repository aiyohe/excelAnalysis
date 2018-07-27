package com.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;

/**
 * @Author: Mr.Zhang
 * @Description: excel 解析 util
 * @Date: 14:58 2018/7/23
 * @Modified By:
 */
public class ExcelAnalysisUtil {
    private static FormulaEvaluator evaluator;

    /**
     * 读取excel 文件
     *
     * @param file
     * @return
     * @throws Exception
     */
    public static JSONArray excelRead(String file, int table) throws Exception {
        Workbook book = null;
        try {
            book = new XSSFWorkbook(new FileInputStream(file));     //解析2003
        } catch (Exception e) {
            book = new HSSFWorkbook(new FileInputStream(file));      //解析2007
        }
        evaluator = book.getCreationHelper().createFormulaEvaluator();
        return getExcelContent(book, table);
    }

    private static JSONArray getExcelContent(Workbook book, int table) throws Exception {
        JSONArray jsar = new JSONArray();
        if (table == 0) {
            table = 1;
        }
        for (int numSheet = 0; numSheet < book.getNumberOfSheets(); numSheet++) {
            Sheet sheet = book.getSheetAt(numSheet);
            if (sheet == null) {   //谨防中间空一行
                continue;
            }
            for (int numRow = table; numRow <= sheet.getLastRowNum(); numRow++) {   //一个row就相当于一个Object
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(numRow);
                if (row == null) {
                    continue;
                }
                jsar.add(getObject(numRow, row));
            }
        }
        return jsar;
    }

    private static JSONObject getObject(int numRow, org.apache.poi.ss.usermodel.Row row) throws Exception {
        JSONObject json = new JSONObject();
        for (int numCell = 0; numCell < row.getLastCellNum(); numCell++) {
            Cell cell = row.getCell(numCell);
            if (cell == null) {
                continue;
            }
            String cellValue = getValue(cell);
            json.put((numRow + 1) + "-" + (numCell + 1), cellValue);
        }
        return json;
    }

    private static String getValue(Cell cell) {
        if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                return sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
            }
            return NumberToTextConverter.toText(cell.getNumericCellValue());
        } else if (cell.getCellType() == cell.CELL_TYPE_STRING) {
            return String.valueOf(cell.getStringCellValue());
        } else if (cell.getCellType() == cell.CELL_TYPE_ERROR) {
            return null;
        } else if (cell.getCellType() == cell.CELL_TYPE_FORMULA) {
            return String.valueOf(evaluator.evaluate(cell).getNumberValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    public static void main(String[] args) {
        JSONArray array = null;
        try {
            array = excelRead("E:/cs/cs.xlsx", 1);
        } catch (Exception e) {
            e.printStackTrace();
        }
        for (int i = 0; i < array.size(); i++) {
            System.out.println("数据-->" + array.get(i).toString());
        }
    }
}
