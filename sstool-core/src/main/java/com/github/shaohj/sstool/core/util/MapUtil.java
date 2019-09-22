package com.github.shaohj.sstool.core.util;

import java.util.Map;

/**
 * 编  号：
 * 名  称：MapUtil
 * 描  述：
 * 完成日期：2019/6/29 10:28
 * @author：felix.shao
 */
public class MapUtil {

    /** hashmap默认加载因子 */
    public static final float DEFAULT_LOAD_FACTOR = 0.75f;

    public static void main(String[] args) {
        System.out.println(calMapSize(7));
    }

    /**
     * 计算Map的默认容量
     * @param size eq:10, return 10 * DEFAULT_LOAD_FACTOR
     * @return
     */
    public static int calMapCapacity(int size){
        return ((int)(size * DEFAULT_LOAD_FACTOR));
    }

    public static int calMapSize(int capacity){
        return ((int)(capacity / DEFAULT_LOAD_FACTOR) + 1);
    }


    /**
     * map是否为空
     * @param map
     * @return
     */
    public static boolean isEmpty(Map<?, ?> map) {
        return null == map || map.isEmpty();
    }

}
