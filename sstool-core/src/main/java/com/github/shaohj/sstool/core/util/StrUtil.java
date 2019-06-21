package com.github.shaohj.sstool.core.util;

import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 编  号：
 * 名  称：StrUtil
 * 描  述：字符串工具类
 * 完成日期：2019/6/18 21:04
 * @author：felix.shao
 */
public class StrUtil {

    /**
     * 判断字符是否为空
     * @param str
     * @return
     */
    public static boolean isEmpty(CharSequence str) {
        return str == null || str.length() == 0;
    }

    /**
     * 获取文件类型后缀
     * @param str eq: a.txt
     * @return String eq: txt
     */
    public static String getStringSuffix(String str) {
        if (isEmpty(str)) {
            return "";
        }
        int idx = str.lastIndexOf(".");
        return idx > 0 ? str.substring(idx + 1) : "";
    }

    /**
     * 字符串模板工具类,只支持${}单个的表达式，不支持嵌套表达式
     * @param content 示例：hello ${name}, 1 2 3 4 5 ${six} 7, again ${name}.
     * @param map  map.put("name", "java"); map.put("six", "6");
     * @return String eq:hello java, 1 2 3 4 5 6 7, again java.
     */
    public static String renderString(String content, Map<String, Object> map){
        Set<Map.Entry<String, Object>> sets = map.entrySet();
        for(Map.Entry<String, Object> entry : sets) {
            String regex = "\\$\\{" + entry.getKey() + "\\}";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(content);
            content = matcher.replaceAll(entry.getValue().toString());
        }
        return content;
    }

}
