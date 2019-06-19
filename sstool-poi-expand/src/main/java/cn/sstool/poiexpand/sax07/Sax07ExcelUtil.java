package cn.sstool.poiexpand.sax07;

import cn.sstool.poiexpand.common.bean.read.ReadSheetData;
import cn.sstool.poiexpand.common.bean.write.WriteSheetData;
import cn.sstool.poiexpand.common.util.write.SaxWriteUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：Sax07ExcelUtil
 * 描  述：
 * 完成日期：2019/01/28 23:13
 *
 * @author：felix.shao
 */
@Slf4j
public class Sax07ExcelUtil {

    /**
     * 根据excel模板导出excel，没有大数据分页标签
     * @param tempFileName 模板文件名
     * @param params javabean参数
     * @param out 输出文件流
     */
    public static void export(String tempFileName, Map<String, Object> params, OutputStream out) {
        List<Sax07ExcelPageWriteService> sax07ExcelPageWriteServices = new ArrayList<>(0);
        export(tempFileName, params, out, sax07ExcelPageWriteServices);
    }

    /**
     * 根据excel模板导出excel，有一个大数据分页标签
     * @param tempFileName 模板文件名
     * @param params javabean参数
     * @param out 输出文件流
     * @param sax07ExcelPageWriteService 一个大数据分页标签处理Service
     */
    public static void export(String tempFileName, Map<String, Object> params, OutputStream out, Sax07ExcelPageWriteService sax07ExcelPageWriteService) {
        List<Sax07ExcelPageWriteService> sax07ExcelPageWriteServices = new ArrayList<>(1);
        sax07ExcelPageWriteServices.add(sax07ExcelPageWriteService);
        export(tempFileName, params, out, sax07ExcelPageWriteServices);
    }

    /**
     * 根据excel模板导出excel，有多个大数据分页标签
     * @param tempFileName 模板文件名
     * @param params javabean参数
     * @param out 输出文件流
     * @param sax07ExcelPageWriteServices 多个大数据分页标签处理Services
     */
    public static void export(String tempFileName, Map<String, Object> params, OutputStream out, List<Sax07ExcelPageWriteService> sax07ExcelPageWriteServices) {
        // 打开工作簿 并进行初始化
        XSSFWorkbook readWb = Sax07ExcelWorkbookUtil.openWorkbookByProPath(tempFileName);

        // 读取模板工作簿内容
        List<ReadSheetData> readReadSheetData = Sax07ExcelReadUtil.readSheetData(readWb);

        // 写入模板内容
        List<WriteSheetData> writeSheetDatas = SaxWriteUtil.parseSheetData(readReadSheetData);

        //创建缓存的输出文件工作簿
        SXSSFWorkbook writeWb = new SXSSFWorkbook(100);

        Sax07ExcelWriteUtil.writeSheetData(writeWb, params, writeSheetDatas, sax07ExcelPageWriteServices);

        try {
            //输出文件
            writeWb.write(out);
        } catch (IOException e) {
            log.info("导出excel时错误", e);
        } finally {
            //清理导出的临时缓存文件
            writeWb.dispose();
        }

    }

}
