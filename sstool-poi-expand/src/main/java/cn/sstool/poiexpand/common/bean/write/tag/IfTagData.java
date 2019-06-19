package cn.sstool.poiexpand.common.bean.write.tag;

import cn.sstool.poiexpand.common.consts.SaxExcelConst;
import cn.sstool.poiexpand.common.consts.TagEnum;
import cn.sstool.core.util.ExprUtil;
import cn.sstool.core.util.StrUtil;
import cn.sstool.poiexpand.common.bean.write.WriteSheetData;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Map;

/**
 * 编  号：
 * 名  称：IfTagData
 * 描  述：if表达式标签，表达书为真，才会解析标签内的内容
 * 完成日期：2019/6/19 23:47
 * @author：felix.shao
 */
public class IfTagData extends TagData{

    private boolean isExprTrue(Map<String, Object> params){
        String exprStr = getRealExpr();
        Object exprValue = ExprUtil.getExprStrValue(params, exprStr);
        boolean isFlag = null == exprValue || StrUtil.isEmpty(exprValue.toString()) ? false : true;
        return isFlag;
    };

    @Override
    public String getRealExpr() {
        return null != value ?
                String.valueOf(value).replace(SaxExcelConst.TAG_KEY + TagEnum.IF_TAG.getKey(), "").trim()
                : "";
    }

    @Override
    public void writeTagData(Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData,
                             Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache) {
        if(isExprTrue(params)){
            childTagDatas.stream().forEach(childTagData -> {
                childTagData.writeTagData(writeWb, writeSheet, writeSheetData, params, writeCellStyleCache);
            });
        }
    }

}
