package com.github.shaohj.sstool.poiexpand.sax07.service;

import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.TagData;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Map;

/**
 * 编  号：
 * 名  称：Sax07ExcelPageWriteService
 * 描  述：定义大数据分页导出的service规范
 * 完成日期：2019/02/05 13:20
 *
 * @author：felix.shao
 */
@Data
public abstract class Sax07ExcelPageWriteService {

    /** 需要写入的excel模板标签块 */
    protected TagData tagData;

    /**  分页的表达式值，和excel模板的表达式比较，找到service对应处理部分 */
    protected String exprVal;

    /**  write Workbook*/
    protected Workbook writeWb;

    /**  write SXSSFSheet */
    protected SXSSFSheet writeSheet;

    /**  write WriteSheetData */
    protected WriteSheetData writeSheetData;

    /**  write样式缓存，创建多样式导致导出文件过大 */
    protected Map<String, CellStyle> writeCellStyleCache;

    public void init(TagData tagData, Workbook writeWb,
                     SXSSFSheet writeSheet, WriteSheetData writeSheetData, Map<String, CellStyle> writeCellStyleCache){
        this.tagData = tagData;
        this.writeWb = writeWb;
        this.writeSheet = writeSheet;
        this.writeSheetData = writeSheetData;
        this.writeCellStyleCache = writeCellStyleCache;
    }
    /**
     * 分页查询出data，并调用tagData.writeTagData(writeWb, writeSheet, writeSheetData, pageParams, writeCellStyleCache);将分页的数据写入Excel
     */
    public abstract void pageWriteData();

}
