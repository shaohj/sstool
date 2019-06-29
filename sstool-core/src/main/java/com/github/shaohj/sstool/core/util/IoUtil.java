package com.github.shaohj.sstool.core.util;

/**
 * 编  号：
 * 名  称：IoUtil
 * 描  述：IO流工具类
 * 完成日期：2019/6/18 21:05
 * @author：felix.shao
 */
public class IoUtil {

	/**
	 * 静默关闭
	 *
	 * @param closeable 被关闭的对象
	 */
	public static void close(AutoCloseable closeable) {
		if (null != closeable) {
			try {
				closeable.close();
			} catch (Exception e) {
				// 静默关闭
			}
		}
	}

}
