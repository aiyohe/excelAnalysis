package com.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

/**
 * @Author: Mr.Zhang
 * @Description: 减少冗余代码
 * @Date: 16:47 2018/7/25
 * @Modified By:
 */
public class AnalysisUtil {
    /**
     * 根据中文字符去匹配对应的 数据
     * 如 单位名称  报告期
     *
     * @param jsonArray
     * @return
     */
    public static String getSpecificValue(JSONArray jsonArray,String specific) {

        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject json = (JSONObject) jsonArray.get(i);
            String dwmc = null;
            String key;

            for (String jsonKey : json.keySet()) {
                String value = (String) json.get(jsonKey);
                if (specific.equals(value)) {
                    key = jsonKey;
                    String[] keyS = key.split("-");
                    dwmc = (String) json.get(keyS[0] + "-" + (Integer.valueOf(keyS[1]) + 1));
                    break;
                }
            }
            return dwmc;
        }
        return  null;
    }
    /**
     * 获取下一个坐标数据
     *
     * @param key
     * @param json
     * @return
     */
    public static String getNextValue(String key, JSONObject json) {
        String[] keyS = key.split("-");
        String value = (String) json.get(keyS[0] + "-" + (Integer.valueOf(keyS[1]) + 1));
        return value;
    }
}