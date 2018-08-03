package com.excel.analysis;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.excel.util.AnalysisUtil;
import com.excel.util.ExcelAnalysisUtil;

/**
 * @Author: Mr.Zhang
 * @Description: 成品油购进、销售及库存监测日报模板
 * @Date: 16:50 2018/7/24
 * @Modified By:
 */
public class Zywz007Day {
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
                        jsa.put("B070101", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报告期：")) {
                        jsa.put("B070102", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    //数据
                    if (value.contains("代码")) {
                        //获取 列数
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("单位负责人：")) {
                        jsa.put("B070103", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("填表人：")) {
                        jsa.put("B070104", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("联系电话：")) {
                        jsa.put("B070105", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报出日期：")) {
                        jsa.put("B070106", AnalysisUtil.getNextValue(jsonKey, json));
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
                        int k = 3;
                        for (int j = Integer.valueOf(x) + 1; j <= json.size(); j++) {
                            jsa.put(code + "-B" + k, json.get(y + "-" + j));
                            k++;
                        }
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
        JSONArray array = excelAnalysis("E:/cs/成品油购进、销售及库存监测表.xlsx");
        for (int i = 0; i < array.size(); i++) {
            System.out.println("数据-->" + array.get(i).toString());
        }
    }
}
