package com.github.shaohj.sstool.poiexpand.common.bean.write.tag;

import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.read.RowData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.util.write.TagUtil;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：ConstTagData
 * 描  述：常量标签块
 * 完成日期：2019/6/19 23:37
 * @author：felix.shao
 */
@Data
@EqualsAndHashCode(callSuper = false)
@NoArgsConstructor
public class ConstTagData extends TagData{

    /** 常量块的内容需要写入，和读取excel模板行数据一致 */
    protected List<RowData> readRowData;

    public ConstTagData(List<RowData> readRowData) {
        this.readRowData = readRowData;
    }

    public void addRowData(RowData rowData){
        if(null == readRowData){
            //一般标签内的excel行数不会太多，设置为4吧
            readRowData = new ArrayList<>(4);
        }
        readRowData.add(rowData);
    }

    @Override
    public String getRealExpr() {
        return String.valueOf(value);
    }

    @Override
    public void writeTagData(Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData,
                             Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache) {
        int curWriteRowNum = writeSheetData.getCurWriteRowNum();
        TagUtil.writeTagData(writeWb, writeSheet, writeSheetData, readRowData, params, writeCellStyleCache);

        if(!MapUtil.isEmpty(allCellRangeAddress)){
            allCellRangeAddress.forEach((idx, craddr) -> {
                //起始行,结束行,起始列,结束列
                CellRangeAddress callRangeAddressInfo = new CellRangeAddress(curWriteRowNum + craddr.getRelaStartRow(), curWriteRowNum + craddr.getRelaEndRow(),
                        craddr.getStartCol(),craddr.getEndCol());
                //addMergedRegionUnsafe比addMergedRegion方法少异常检查，大大优化导出时间
                writeSheet.addMergedRegionUnsafe(callRangeAddressInfo);
            });
        }

    }

}
