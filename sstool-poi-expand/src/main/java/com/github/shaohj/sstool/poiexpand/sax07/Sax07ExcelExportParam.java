package com.github.shaohj.sstool.poiexpand.sax07;

import com.github.shaohj.sstool.core.util.StrUtil;
import com.github.shaohj.sstool.poiexpand.sax07.service.Sax07ExcelPageWriteService;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：Sax07ExcelExportParam
 * 描  述：
 * 完成日期：2019/06/29 12:00
 *
 * @author：felix.shao
 */
@Builder(toBuilder = true)
@Getter
@Setter
public class Sax07ExcelExportParam {

    /** 模板文件路径是否为class路径.true:是;false:否 */
    private boolean tempIsClassPath;

    /** 模板文件名称 */
    private String tempFileName;

    /** 导出java参数 */
    private Map<String, Object> params;

    /** 输出文件流 */
    private OutputStream outputStream;

    /** 分页导出处理service，支持多页分页导出 */
    private List<Sax07ExcelPageWriteService> sax07ExcelPageWriteServices;

    /**
     * 验证导出参数设置是否合理
     * @param param
     */
    public static void valid(Sax07ExcelExportParam param){
        if (StrUtil.isEmpty(param.getTempFileName())){
            throw new IllegalArgumentException("tempFileName can not be empty");
        }

        if (null == param.getParams()){
            throw new IllegalArgumentException("param can not be empty");
        }

        if (null == param.getOutputStream()){
            throw new IllegalArgumentException("outputStream  can not be empty");
        }
    }

}
