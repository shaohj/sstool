package com.github.shaohj.sstool.wordexpand.util;

import com.github.shaohj.sstool.wordexpand.freemarker.FreemarkerUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

/**
 * 真实路径导出测试
 * 编  号：
 * 名  称：FmRealPathManagerTest
 * 描  述：
 * 完成日期：2018/11/18 22:12
 * @author：felix.shao
 */
@Slf4j
public class WordUtilTest {

	/** word存放的真实目录 */
	public static final String ROOT_DOCX_TEMPFILE_FOLDER_URL = "E:\\temp\\word_template\\docx\\";

	/** word的ftl模板文件存放目录 */
	public static final String ROOT_FREEMARKER_TEMPFILE_FOLDERURL = "E:\\temp\\word_template\\ftl";

	public static final String EXPORT_WORD_PATH = "E:\\temp\\wordExport\\";

	/**
	 * 根据ftl模板文件输出string
	 */
	@Test
	public void parseToStringTest(){
		Map<String, Object> params = new HashMap<String, Object>(8);
		params.put("userName", "张三");
		params.put("sex", "男");
		params.put("address", "深圳福田区");

		DocxExportParam docxExportParam = DocxExportParam.builder()
				.freemarkerVersion("2.3.28")
				.rootFreemarkerTempFileFolderUrl(ROOT_FREEMARKER_TEMPFILE_FOLDERURL)
				.tempFileUrl("testtmp.ftl")
				.params(params)
				.build();

		try {
			log.info("parseResult=\n{}", FreemarkerUtil.parseTFToString(docxExportParam));
		} catch (Exception e) {
			log.error("{}", e);
		}
	}

	/**
	 * 导出docx
	 */
	@Test
	public void reportDocx(){
		Map<String, Object> params = new HashMap<String, Object>(8);
		params.put("startNo", "你好,张三");
		params.put("numNo", "10086");
		params.put("title", "freemarker word 模板导出docx");
		params.put("endMsg", "导出结束啦");

		String toPath = EXPORT_WORD_PATH + "中文名导出测试.docx";

		DocxExportParam docxExportParam = DocxExportParam.builder()
				.freemarkerVersion("2.3.28")
				.rootDocxTempFileFolderUrl(ROOT_DOCX_TEMPFILE_FOLDER_URL)
				.rootFreemarkerTempFileFolderUrl(ROOT_FREEMARKER_TEMPFILE_FOLDERURL)
				.tempFileUrl("\\中文名导出.docx")
				.params(params)
				.toFilePath(toPath)
				.build();

		//导出docx,file导出,String导出时文件太大会溢出
		try {
			//导出docx
			WordUtils.parseDocx(docxExportParam);
		} catch (Exception e) {
			log.error("{}", e);
		}
	}

}
