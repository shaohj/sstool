package com.github.shaohj.sstool.poiexpand.common.bean.write.tag;

import com.github.shaohj.sstool.poiexpand.common.bean.write.MergeRegionParam;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：TagData
 * 描  述：excel模板标签块，如常量，if，foreach，pageforeach，each块等
 * 完成日期：2019/6/19 23:38
 * @author：felix.shao
 */
@Data
public abstract class TagData {

    /** 表达式值，常量为常量，表达式如#pageforeach detail in ${list}等 */
    protected Object value;

    /** 写入sheet对应的所有合并单元格区域参数,要求是合并单元格必须再每个写入块tagdata中，此逻辑先不作验证，配置模板时注意下就可以了 */
    protected Map<String, MergeRegionParam> allCellRangeAddress;

    /** 标签块中还有标签块，如if标签块中嵌套多层块 */
    protected List<TagData> childTagDatas = new ArrayList<>(4);

    /** 获取真实的表达式值，每个标签的表达式处理逻辑不一样 */
    public abstract String getRealExpr();

    /** 根据标签解析tag，每个标签的tag写入逻辑不一样 */
    public abstract void writeTagData(Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData,
                                      Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache);

}
