package com.github.shaohj.sstool.core.util;

import java.util.Collection;
import java.util.Map;

/**
 * 编  号：
 * 名  称：EmptyUtil
 * 描  述：empty工具类，设计目的是仅sstool使用。
 * 完成日期：2019/06/18 21:06
 *
 * @author：felix.shao
 */
public class EmptyUtil {

    /**
     * 数组是否为空
     * @param array
     * @return
     */
    public static boolean isEmpty(final int... array) {
        return array == null || array.length == 0;
    }

    /**
     * 集合是否为空
     * @param collection
     * @return
     */
    public static boolean isEmpty(Collection<?> collection) {
        return collection == null || collection.isEmpty();
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
