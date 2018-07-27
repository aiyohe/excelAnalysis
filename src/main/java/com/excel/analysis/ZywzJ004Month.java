package com.excel.analysis;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.excel.util.AnalysisUtil;
import com.excel.util.ExcelAnalysisUtil;

/**
 * @Author: Mr.Zhang
 * @Description: 货物发送到达完成情况（月报）
 * @Date: 9:50 2018/7/25
 * @Modified By:
 */

public class ZywzJ004Month {
    /**
     * @param filePath 文件路径
     * @return
     */
    public static JSONArray excelAnalysis(String filePath) {
        JSONArray jsar = new JSONArray();
        String x = null;
        String y = null;
        try {
            JSONArray jsonArray = ExcelAnalysisUtil.excelRead(filePath, 1);
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject json = (JSONObject) jsonArray.get(i);
                JSONObject jsa = new JSONObject();
                for (String jsonKey : json.keySet()) {
                    String value = (String) json.get(jsonKey);
                    if (value.contains("单位名称：")) {
                        jsa.put("B010102", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报告期：")) {
                        jsa.put("B010103", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    //数据
                    if (value.contains("序号")) {
                        //获取 列数
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("单位负责人：")) {
                        jsa.put("B010104", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("填表人：")) {
                        jsa.put("B010105", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("联系电话：")) {
                        jsa.put("B010106", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报出日期：")) {
                        jsa.put("B010107", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    y = jsonKey.split("-")[0];
                }
                //处理数据
                if (x != null) {
                    Integer code;
                    try {
                        code = Integer.valueOf((String) json.get(y + "-" + x));
                    } catch (Exception e) {
                        e.getMessage();
                        code = 0;
                    }
                    if (code != 0) {
                        int xx = Integer.valueOf(x);
                        jsa.put(code + "-B3", json.get(y + "-" + (xx + 2)));
                        jsa.put(code + "-B4", json.get(y + "-" + (xx + 3)));
                        jsa.put(code + "-B9", json.get(y + "-" + (xx + 4)));
                        jsa.put(code + "-B10", json.get(y + "-" + (xx + 5)));
                    }
                }
                if (jsa.size() > 0) {
                    jsar.add(jsa);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return jsar;
    }

    public static void main(String[] args) {
        JSONArray array = excelAnalysis("E:/cs/货物发送到达完成情况（月报）.xlsx");
        for (int i = 0; i < array.size(); i++) {
            System.out.println("数据-->" + array.get(i).toString());
        }
    }
}
