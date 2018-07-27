package com.excel.analysis;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.excel.util.AnalysisUtil;
import com.excel.util.ExcelAnalysisUtil;

/**
 * @Author: Mr.Zhang
 * @Description: 重点监测单位成品油分区域加油站销售监测月报表
 * @Date: 9:46 2018/7/25
 * @Modified By:
 */
public class ZywzCPY002 {
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
                int code = 0;
                for (String jsonKey : json.keySet()) {
                    String value = (String) json.get(jsonKey);
                    if (value.contains("单位名称：")) {
                        jsa.put("A005", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报告期：")) {
                        jsa.put("A006", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    //获取用户数
                    if (value.contains("主要用户")) {
                        code = 9;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("销售量")) {
                        //获取 列数
                        code = 1;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if ("环比增加".equals(value)) {
                        //获取 列数
                        code = 2;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("环比增加％")) {
                        //获取 列数
                        code = 3;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if ("同比增加".equals(value)) {
                        //获取 列数
                        code = 4;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("同比增加％")) {
                        //获取 列数
                        code = 5;
                        x = jsonKey.split("-")[1];
                        continue;
                    }
                    if (value.contains("月度简析")) {
                        jsa.put("YDJX", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("单位负责人：")) {
                        jsa.put("A003", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("填表人：")) {
                        jsa.put("A001", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("联系电话：")) {
                        jsa.put("A002", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    if (value.contains("报出日期：")) {
                        jsa.put("A004", AnalysisUtil.getNextValue(jsonKey, json));
                    }
                    y = jsonKey.split("-")[0];
                }
                //获取销售量
                if (code != 0 && code != 9) {
                    int num = json.size();
                    int xx = Integer.valueOf(x) + 1;
                    for (int j = 1; j <= num; j++) {
                        String value = (String) json.get(y + "-" + xx);
                        if (value != null) {
                            jsa.put(code + "-B" + j, value);
                        }
                        xx++;
                    }
                }
                //获取多少用户
                if (code == 9) {
                    int num = 1;
                    for (int j = 1; j <= json.size(); j++) {
                        String va = (String) json.get(y + "-" + j);
                        try {
                            jsa.put("" + num, Integer.valueOf(va));
                        } catch (Exception e) {
                            continue;
                        }
                        num++;
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
        JSONArray array = excelAnalysis("E:/cs/重点监测单位成品油分区域加油站销售监测月报表.xlsx");
        for (int i = 0; i < array.size(); i++) {
            System.out.println("数据-->" + array.get(i).toString());
        }
    }
}
