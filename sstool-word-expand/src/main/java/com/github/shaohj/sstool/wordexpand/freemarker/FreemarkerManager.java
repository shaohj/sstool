package com.github.shaohj.sstool.wordexpand.freemarker;

import com.github.shaohj.sstool.wordexpand.util.DocxExportParam;
import freemarker.cache.FileTemplateLoader;
import freemarker.cache.TemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.TemplateExceptionHandler;
import freemarker.template.Version;

import java.io.File;
import java.io.IOException;

/**
 * 编  号：
 * 名  称：FreeMarkerManager
 * 描  述：
 * 完成日期：2018/11/18 11:35
 * @author：felix.shao
 */
public class FreemarkerManager {

	/** freemark模板文件在本地真实路径下的配置 */
	private static volatile Configuration realPathCfg;

	/**
	 * 指定路径加载,替换模板文件后不会重启服务器,动态生成了模板路径，因此没考虑封装相对路径方式
	 * @return
	 * @author felix.shao
	 * @since 2018/11/18 21:19
	 */
	public static Configuration singleRealPathCfg(DocxExportParam docxExportParam) throws IOException {
		if(null == realPathCfg){
			synchronized (FreemarkerManager.class){
				if(null == realPathCfg){
					Version version = new Version(docxExportParam.getFreemarkerVersion());

					//定义加载真实路径下的模板文件
					TemplateLoader tempLoader = new FileTemplateLoader(new File(docxExportParam.getRootFreemarkerTempFileFolderUrl()));

					realPathCfg = new Configuration(version);
					realPathCfg.setTemplateLoader(tempLoader);
					realPathCfg.setNumberFormat("0");
					//设置对象的包装器
					realPathCfg.setObjectWrapper(new DefaultObjectWrapper(version));
					//设置异常处理器	${a.b.c.d}即使没有属性也不会出错
					realPathCfg.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
					realPathCfg.setDefaultEncoding("UTF-8");
				}
			}
		}
		return realPathCfg;
	}

}
