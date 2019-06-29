package com.github.shaohj.sstool.poiexpand.common.bean.write.tag;

import com.github.shaohj.sstool.poiexpand.common.consts.SaxExcelConst;
import com.github.shaohj.sstool.poiexpand.common.consts.TagEnum;
import com.github.shaohj.sstool.core.util.ExprUtil;
import com.github.shaohj.sstool.core.util.StrUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Iterator;
import java.util.Map;
import java.util.StringTokenizer;

/**
 * 编  号：
 * 名  称：ForeachTagData
 * 描  述：循环一次性导出list
 * 完成日期：2019/6/19 23:47
 * @author：felix.shao
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class ForeachTagData extends TagData {

    private String iteratorObjKey;

    private String iteratorListKey;

    @Override
    public String getRealExpr() {
        return null != value ?
                String.valueOf(value).replace(SaxExcelConst.TAG_KEY + TagEnum.FOREACH_TAG.getKey(), "").trim()
                : "";
    }

    @Override
    public void writeTagData(Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData, Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache) {
        initExpr();
        Object iteratorList = ExprUtil.getExprStrValue(params, iteratorListKey);
        if(null == iteratorListKey){
            return;
        }
        Iterator iterator = ExprUtil.getIterator(iteratorList);

        while (iterator.hasNext()) {
            Object iteratorObj = iterator.next();
            params.put(iteratorObjKey, iteratorObj);
            childTagDatas.stream().forEach(childTagData -> {
                childTagData.writeTagData(writeWb, writeSheet, writeSheetData, params, writeCellStyleCache);
            });
        }
    }

    private void initExpr(){
        String realExpr = getRealExpr();
        if(StrUtil.isEmpty(realExpr)){
            return;
        }
        StringTokenizer st = new StringTokenizer(realExpr, " ");
        int pos = 0;
        while (st.hasMoreTokens()) {
            String str = st.nextToken();
            if (pos == 0) {
                iteratorObjKey = str;
            } else if (pos == 2) {
                iteratorListKey = str;
            }
            pos++;
        }
    };

}
