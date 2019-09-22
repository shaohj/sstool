package com.github.shaohj.sstool.poiexpand.sax07;

import com.github.shaohj.sstool.poiexpand.common.bean.read.ReadSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.util.write.SaxWriteUtil;
import com.github.shaohj.sstool.poiexpand.sax07.template.Sax07ExcelTemplateReader;
import com.github.shaohj.sstool.poiexpand.sax07.template.Sax07ExcelTemplateWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 编  号：
 * 名  称：Sax07ExcelUtil
 * 描  述：
 * 完成日期：2019/6/29 11:12
 * @author：felix.shao
 */
@Slf4j
public class Sax07ExcelUtil {

    /**
     * 根据excel2007模板导出excel
     * @param param 导出参数
     */
    public static void export(Sax07ExcelExportParam param) {
        Sax07ExcelExportParam.valid(param);
        if (null == param.getSax07ExcelPageWriteServices()){
            param.setSax07ExcelPageWriteServices(new ArrayList<>(0));
        }

        // 打开工作簿 并进行初始化
        XSSFWorkbook readWb = Sax07ExcelTemplateReader.openWorkbook(true, param.getTempFileName());

        // 读取模板工作簿内容
        List<ReadSheetData> readReadSheetData = Sax07ExcelTemplateReader.readSheetData(readWb);

        // 写入模板内容
        List<WriteSheetData> writeSheetDatas = SaxWriteUtil.parseSheetData(readReadSheetData);

        //创建缓存的输出文件工作簿
        int rowAccessWindowSize = 0 == param.getRowAccessWindowSize() ? 1000 : param.getRowAccessWindowSize();
        SXSSFWorkbook writeWb = new SXSSFWorkbook(rowAccessWindowSize);

        try {
            Sax07ExcelTemplateWriter.writeSheetData(writeWb, param.getParams(), writeSheetDatas, param.getSax07ExcelPageWriteServices());
            //输出文件
            writeWb.write(param.getOutputStream());
        } catch (IOException e) {
            log.info("导出excel时错误", e);
        } finally {
            //清理导出的临时缓存文件
            writeWb.dispose();
        }

    }

}
