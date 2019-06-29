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

    /**
     * 计算Map的默认容量
     * @param capacity
     * @return
     */
    public static int calMapCapacity(int capacity){
        return ((int)(capacity * DEFAULT_LOAD_FACTOR) + 1);
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
