package com.github.shaohj.sstool.poiexpand.sax07.template;

import com.github.shaohj.sstool.core.util.EmptyUtil;
import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.PageForeachTagData;
import com.github.shaohj.sstool.poiexpand.sax07.service.Sax07ExcelPageWriteService;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：Sax07ExcelTemplateWriter
 * 描  述：
 * 完成日期：2019/6/29 11:12
 * @author：felix.shao
 */
public class Sax07ExcelTemplateWriter {

    /**
     * 写入sheet数据，有多个大数据分页标签
     * @param writeWb
     * @param params
     * @param writeSheetDatas
     * @param sax07ExcelPageWriteServices
     */
    public static void writeSheetData(SXSSFWorkbook writeWb, Map<String, Object> params, List<WriteSheetData> writeSheetDatas, List<Sax07ExcelPageWriteService> sax07ExcelPageWriteServices){
        if(EmptyUtil.isEmpty(writeSheetDatas)){
            return;
        }
        writeSheetDatas.stream().forEach(writeSheetData -> {
            //创建sheet
            SXSSFSheet writeSheet =  writeWb.createSheet(writeSheetData.getSheetName());

            /** 设置列宽为模板文件的列宽 */
            if(!EmptyUtil.isEmpty(writeSheetData.getCellWidths())){
                for (int i = 0; i < writeSheetData.getCellWidths().length; i++) {
                    writeSheet.setColumnWidth(i, writeSheetData.getCellWidths()[i]);
                }
            }

            if(MapUtil.isEmpty(writeSheetData.getWriteTagDatas())){
                return;
            }
            // 样式需要做缓存特殊处理，以sheet为单位作缓存处理，定义在此保证线程安全
            final Map<String, CellStyle> writeCellStyleCache = new HashMap<>();
            writeSheetData.getWriteTagDatas().forEach((readRowNum, tagData) -> {
                if(tagData instanceof PageForeachTagData){
                    // 大数据导出service
                    if(null != tagData.getValue() && !EmptyUtil.isEmpty(sax07ExcelPageWriteServices)){
                        sax07ExcelPageWriteServices.stream()
                                .filter(sax07ExcelPageWriteService -> sax07ExcelPageWriteService.getExprVal().equalsIgnoreCase(String.valueOf(tagData.getValue())))
                                .findFirst().ifPresent(sax07ExcelPageWriteService -> {
                            sax07ExcelPageWriteService.init(tagData, writeWb, writeSheet, writeSheetData, writeCellStyleCache);
                            sax07ExcelPageWriteService.pageWriteData();
                        });
                    }
                } else {
                    tagData.writeTagData(writeWb, writeSheet, writeSheetData, params, writeCellStyleCache);
                }
            });
        });
    }

}
