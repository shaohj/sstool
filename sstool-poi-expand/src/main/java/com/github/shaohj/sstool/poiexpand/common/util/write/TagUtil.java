package com.github.shaohj.sstool.poiexpand.common.util.write;

import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.read.RowData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.MergeRegionParam;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.*;
import com.github.shaohj.sstool.poiexpand.common.consts.TagEnum;
import com.github.shaohj.sstool.poiexpand.common.exception.PoiExpandException;
import com.github.shaohj.sstool.poiexpand.common.util.ExcelCommonUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：TagUtil
 * 描  述：
 * 完成日期：2019/6/19 23:53
 * @author：felix.shao
 */
public class TagUtil {

    public static TagEnum getTagEnum(RowData rowData){
        String expr = SaxWriteUtil.getFirstCellValueStr(rowData);
        return TagEnum.getTagEnum(expr);
    };

    public static TagData getTagData(TagEnum tagEnum){
        TagData tagData = null;
        switch (tagEnum){
            case IF_TAG:
                tagData = new IfTagData();
                break;
            case FOREACH_TAG:
                tagData = new ForeachTagData();
                break;
            case BIGFOREACH_TAG:
                tagData = new PageForeachTagData();
                break;
            case EACH_TAG:
                tagData = new EachTagData();
                break;
            case CONST_TAG:
                tagData = new ConstTagData();
                break;
            default:
                throw new PoiExpandException("不支持的标签格式");
        }
        return tagData;
    }

    /**
     * 获取常量标签的最后一个行号
     * @param rowNumStart
     * @param rowNumEnd
     * @param rowDatas
     * @return int 常量标签行号
     */
    public static int getConstTagEndNum(int rowNumStart, int rowNumEnd, Map<String, RowData> rowDatas){
        if(rowNumStart < 0 || rowNumStart > rowNumEnd || rowNumEnd >= rowDatas.size()){
            return rowNumEnd;
        }

        int curRowNum = rowNumStart;

        while(curRowNum <= rowNumEnd && rowNumEnd < rowDatas.size()){
            RowData rowData = rowDatas.get(String.valueOf(curRowNum));
            TagEnum tagEnum = getTagEnum(rowData);

            if(tagEnum != TagEnum.CONST_TAG){
                return curRowNum - 1;
            } else if(isEndTag(rowData)){
                return curRowNum - 1;
            }
            curRowNum ++;
        }
        return rowNumEnd;
    }

    public static int getTagEndNum(int rowNumStart, int rowNumEnd, Map<String, RowData> rowDatas){
        if(rowNumStart < 0 || rowNumStart > rowNumEnd || rowNumEnd >= rowDatas.size()){
            return -1;
        }

        int curRowNum = rowNumStart;

        // tag存在嵌套
        int tagNum = 0;
        while(curRowNum <= rowNumEnd && rowNumEnd < rowDatas.size()){
            RowData rowData = rowDatas.get(String.valueOf(curRowNum));
            TagEnum tagEnum = getTagEnum(rowData);
            if(tagEnum.isHasEndTag()){
                // 嵌套标签且有结束标签
                tagNum ++;
            } else {
                boolean isEnd = isEndTag(rowData);
                if (isEnd) {
                    if(0 == tagNum){
                        return curRowNum;
                    } else {
                        tagNum --;
                    }
                }
            }
            curRowNum ++;
        }
        return -1;
    }

    public static void writeTagData(Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData,
                                    List<RowData> readRowData, Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache){
        readRowData.stream().forEach(rowData -> {
            final Row writeRow = writeSheet.createRow(writeSheetData.getCurWriteRowNumAndIncrement());
            writeRow.setHeight(rowData.getHeight());
            writeRow.setHeightInPoints(rowData.getHeightInPoints());
            if(MapUtil.isEmpty(rowData.getCellDatas())){
                return;
            }
            rowData.getCellDatas().forEach((readCellNum, cellData) -> SaxWriteUtil.writeCellData(writeWb, writeRow, rowData.getRowNum(), cellData, params, writeCellStyleCache));
        });
    }

    public static void writeTagMergeRegion(Map<String, MergeRegionParam> allCellRangeAddress, int curWriteRowNum,
                                           Workbook writeWb, SXSSFSheet writeSheet, WriteSheetData writeSheetData){
        if(!MapUtil.isEmpty(allCellRangeAddress)){
            allCellRangeAddress.forEach((idx, craddr) -> {
                //起始行,结束行,起始列,结束列
                CellRangeAddress region = new CellRangeAddress(curWriteRowNum + craddr.getRelaStartRow(), curWriteRowNum + craddr.getRelaEndRow(),
                        craddr.getStartCol(),craddr.getEndCol());
                //addMergedRegionUnsafe比addMergedRegion方法少异常检查，大大优化导出时间
//                writeSheet.addMergedRegionUnsafe(region);
                // 使用下面方式在原有基础上去掉了返回值，原有返回值中有锁等一些操作，缩短了一些导出时间
                ExcelCommonUtil.addMergeRegion((SXSSFWorkbook)writeWb, writeSheet.getSheetName(), region, writeSheetData.getMergeCellsCount());
                writeSheetData.setMergeCellsCount(writeSheetData.getMergeCellsCount() + 1);
            });
        }
    }

    public static boolean isEndTag(RowData rowData){
        String expr = SaxWriteUtil.getFirstCellValueStr(rowData);
        return TagEnum.isEndTagNum(expr);
    }

}
