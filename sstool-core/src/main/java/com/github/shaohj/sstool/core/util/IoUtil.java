package com.github.shaohj.sstool.core.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

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

	public static void copy(InputStream inp, OutputStream out) throws IOException {
		byte[] buff = new byte[4096];
		int count;
		while ((count = inp.read(buff)) != -1) {
			if (count > 0) {
				out.write(buff, 0, count);
			}
		}
	}

}
