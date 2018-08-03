package com.excel.analysis;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.excel.util.AnalysisUtil;
import com.excel.util.ExcelAnalysisUtil;

/**
 * @Author: Mr.Zhang
 * @Description: 原油购进和成品油、液化石油气生产及库存监测表
 * @Date: 9:32 2018/7/25
 * @Modified By:
 */
public class Zywz006Day {
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
                        jsa.put("B060101", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报告期：")) {
                        jsa.put("B060102", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    //数据
                    if (value.contains("代码")) {
                        //获取 列数
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("单位负责人：")) {
                        jsa.put("B060103", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("填表人：")) {
                        jsa.put("B060104", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("联系电话：")) {
                        jsa.put("B060105", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报出日期：")) {
                        jsa.put("B060106", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    y = jsonKey.split("-")[0];

                }
                //处理数据
                if (x != null) {
                    String code;
                    try {
                        code = (String) json.get(y + "-" + x);
                    } catch (Exception e) {
                        code = null;
                    }
                    if (code != null&&!"".equals(code)) {
                        jsa.put(code + "-B3", json.get(y + "-" + (Integer.valueOf(x) + 1)));
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
        JSONArray array = excelAnalysis("E:/cs/原油购进和成品油、液化石油气生产及库存监测表.xlsx");
        for (int i = 0; i < array.size(); i++) {
            System.out.println("数据-->" + array.get(i).toString());
        }
    }
}
