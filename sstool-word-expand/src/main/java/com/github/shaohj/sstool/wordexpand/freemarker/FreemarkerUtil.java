package com.github.shaohj.sstool.wordexpand.freemarker;

import com.alibaba.fastjson.util.IOUtils;
import com.github.shaohj.sstool.core.util.FileUtils;
import com.github.shaohj.sstool.wordexpand.util.DocxExportParam;
import freemarker.template.Template;
import lombok.extern.slf4j.Slf4j;

import java.io.*;

/**
 * 注意，模板文件路径,模板名称不支持中文路径，即ftl模板文件写为英文即可
 * 编  号：
 * 名  称：FreeMarkerUtil
 * 描  述：
 * 完成日期：2018/11/18 22:02
 * @author：felix.shao
 */
@Slf4j
public class FreemarkerUtil {

	/**
	 * 根据模板文件输出string<br />
	 * 不适用大模板文件场景，否则会数据溢出<br />
	 * @param docxExportParam 导出参数
	 * @rturn
	 */
	public static String parseTFToString(DocxExportParam docxExportParam) {
		StringWriter stringWriter = new StringWriter();
		try {
			Template template = FreemarkerManager.singleRealPathCfg(docxExportParam).getTemplate(docxExportParam.getTempFileUrl());
			template.process(docxExportParam.getParams(), new PrintWriter(stringWriter));
			StringBuffer buffer = stringWriter.getBuffer();
			return new String(buffer.toString().getBytes("UTF-8"),"UTF-8");
		} catch (Exception e) {
			log.error("Error by parseTFToString.{}", e);
		}
		return null;
	}

	/**
	 * 根据模板文件输出新文件<br />
	 * @param docxExportParam 导出参数
	 * @param oldFtlUrl 生成的原始的ftl临时路径,相对于模板文件的路径
	 * @param newFtlUrl 生成的替换了表达式的ftl临时路径
	 * @return
	 */
	public static void parseTFToFile(DocxExportParam docxExportParam, String oldFtlUrl, String newFtlUrl) {
		/** 创建导出父文件夹 */
		FileUtils.createParentFolder(docxExportParam.getToFilePath());

		Writer out = null;
		try {
			Template template = FreemarkerManager.singleRealPathCfg(docxExportParam).getTemplate(oldFtlUrl);
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(newFtlUrl)), "UTF-8"));
            template.process(docxExportParam.getParams(), out);
			out.flush();
		} catch (Exception e) {
			log.error("Error by parseTFToFile.{}", e);
		} finally {
			IOUtils.close(out);
		}
	}
	
}
