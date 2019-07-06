package com.github.shaohj.sstool.wordexpand.util;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.util.Map;

/**
 * 编  号：
 * 名  称：DocxExportParam
 * 描  述：
 * 完成日期：2019/07/06 15:08
 *
 * @author：felix.shao
 */
@Builder(toBuilder = true)
@Getter
@Setter
public class DocxExportParam {

    /** freemarker版本号，默认版本为2.3.28 */
    private String freemarkerVersion;

    /** 根WORD模板文件夹的绝对路径，不使用相对路径的理由是我们会动态生成模板ftl文件操作，因此没放置到classpath中 */
    private String rootDocxTempFileFolderUrl;

    /** 根freemarker模板文件夹的绝对路径，不使用相对路径的理由是我们会动态生成模板ftl文件操作，因此没放置到classpath中 */
    private String rootFreemarkerTempFileFolderUrl;

    /** 模板文件相对模板文件夹根路径的路径，根据模板名称不支持中文路径 */
    private String tempFileUrl;

    /** 导出java参数 */
    private Map<String, Object> params;

    /** 导出后生成文件路径,支持含中文名路径 */
    String toFilePath;

}
