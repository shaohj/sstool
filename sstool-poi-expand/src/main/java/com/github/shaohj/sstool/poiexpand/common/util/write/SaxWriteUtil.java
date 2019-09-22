package com.github.shaohj.sstool.poiexpand.common.util.write;

import com.github.shaohj.sstool.core.util.EmptyUtil;
import com.github.shaohj.sstool.core.util.ExprUtil;
import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.read.CellData;
import com.github.shaohj.sstool.poiexpand.common.bean.read.ReadSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.read.RowData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.MergeRegionParam;
import com.github.shaohj.sstool.poiexpand.common.bean.write.WriteSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.ConstTagData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.EachTagData;
import com.github.shaohj.sstool.poiexpand.common.bean.write.tag.TagData;
import com.github.shaohj.sstool.poiexpand.common.consts.TagEnum;
import com.github.shaohj.sstool.poiexpand.common.exception.PoiExpandException;
import com.github.shaohj.sstool.poiexpand.common.util.ExcelCommonUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * 编  号：
 * 名  称：SaxWriteUtil
 * 描  述：
 * 完成日期：2019/6/19 23:53
 * @author：felix.shao
 */
public class SaxWriteUtil {

    /**
     * 将读取的readSheetDatas转为需要写入的WriteSheetDatas
     * @param readSheetDatas
     * @return
     */
    public static List<WriteSheetData> parseSheetData(List<ReadSheetData> readSheetDatas){
        if(EmptyUtil.isEmpty(readSheetDatas)){
            return new ArrayList<>(0);
        }
        final List<WriteSheetData> writeSheetDatas = new ArrayList<>(readSheetDatas.size());
        readSheetDatas.stream().forEach(readSheetData -> writeSheetDatas.add(parseSheetData(readSheetData)));
        return writeSheetDatas;
    }

    /**
     * 将读取的readSheetData转为需要写入的WriteSheetData
     * @param readSheetData
     * @return
     */
    private static WriteSheetData parseSheetData(ReadSheetData readSheetData){
        WriteSheetData writeSheetData = new WriteSheetData();
        writeSheetData.setSheetNum(readSheetData.getSheetNum());
        writeSheetData.setSheetName(readSheetData.getSheetName());
        writeSheetData.setCellWidths(readSheetData.getCellWidths());
        writeSheetData.setWriteTagDatas(parseRowData(readSheetData, readSheetData.getRowDatas()));

        return writeSheetData;
    }

    /**
     * 将读取的rowDatas判断标签转为写入时的块
     * @param readSheetData 包含全局合并单元格等信息
     * @param rowDatas
     * @return
     */
    public static Map<String, TagData> parseRowData(ReadSheetData readSheetData, Map<String, RowData> rowDatas){
        if(MapUtil.isEmpty(rowDatas)){
            return new HashMap<>(0);
        }

        final Map<String, TagData> writeBlockMap = new LinkedHashMap<>(MapUtil.calMapSize(rowDatas.size()));

        geneTreeWriteBlock(readSheetData,0,rowDatas.size() - 1, rowDatas, null, writeBlockMap);

        return writeBlockMap;
    }

    /**
     * 递归循环遍历读取的RowData，将RowData合并为一个一个的写入块
     *   如第m行是#foreach，第n行是对应的#end，那么第m~n行的代码块合并为一个FOREACH_TAG块
     * @param readSheetData 包含全局合并单元格等信息
     * @param rowNumStart rowNum 块的开始行
     * @param rowNumEnd rowNum 块的结束行
     * @param rowDatas 原始的待写的row数据
     * @param rootTagData 父tagData
     * @param writeBlockMap 转换后的待写的block数据
     * @return
     */
    public static void geneTreeWriteBlock(ReadSheetData readSheetData, int rowNumStart, int rowNumEnd, Map<String, RowData> rowDatas, TagData rootTagData, Map<String, TagData> writeBlockMap){
        if(rowNumStart < 0 || rowNumStart > rowNumEnd || rowNumEnd >= rowDatas.size()){
            return ;
        }
        int writeBlockNum = 0;

        int curRowNum = rowNumStart;
        while(curRowNum <= rowNumEnd && rowNumEnd < rowDatas.size()){
            RowData rowData = rowDatas.get(String.valueOf(curRowNum));

            if(curRowNum == rowNumEnd && null != rootTagData && TagUtil.isEndTag(rowData)){
                // 递归时，若为递归父标签的结束标签，则跳出
                break;
            }
            TagEnum tagEnum = TagUtil.getTagEnum(rowData);
            TagData tagData = TagUtil.getTagData(tagEnum);

            int curRowEndNum = -1;
            switch (tagEnum){
                case IF_TAG:
                case FOREACH_TAG:
                case BIGFOREACH_TAG:
                    tagData.setValue(getFirstCellValueStr(rowData));
                    curRowEndNum = TagUtil.getTagEndNum(curRowNum + 1, rowNumEnd, rowDatas);
                    geneTreeWriteBlock(readSheetData, curRowNum + 1, curRowEndNum, rowDatas, tagData, writeBlockMap);
                    // 动态的合并单元格处理
                    tagData.setAllCellRangeAddress(ExcelCommonUtil.getMergeRegionParam(getFirstCellValueStr(rowData)));

                    curRowNum = curRowEndNum;
                    break;
                case EACH_TAG:
                    ((EachTagData) tagData).setReadRowData(rowData);

                    // 动态的合并单元格处理
                    tagData.setAllCellRangeAddress(ExcelCommonUtil.getMergeRegionParam(getFirstCellValueStr(rowData)));
                    tagData.setValue(getFirstCellValueStr(rowData));
                    break;
                case CONST_TAG:
                    curRowEndNum = TagUtil.getConstTagEndNum(curRowNum + 1, rowNumEnd, rowDatas);
                    List<RowData> constRowDatas = new ArrayList(curRowEndNum - curRowNum);
                    for (int i = curRowNum; i<= curRowEndNum; i++){
                        constRowDatas.add(rowDatas.get(String.valueOf(i)));
                    }
                    ((ConstTagData) tagData).setReadRowData(constRowDatas);

                    // 合并单元格处理
                    if(!MapUtil.isEmpty(readSheetData.getAllCellRangeAddress())){
                        Map<String, MergeRegionParam> allCellRangeAddress = new HashMap<>(2);
                        for(Map.Entry<String, CellRangeAddress> entry: readSheetData.getAllCellRangeAddress().entrySet()){
                            String idx = entry.getKey();
                            CellRangeAddress craddr = entry.getValue();
                            // 判断合并单元格区域是否在本标签内
                            if(craddr.getFirstRow() >= curRowNum && craddr.getLastRow() <= curRowEndNum){
                                MergeRegionParam mergeRegionParam = new MergeRegionParam();
                                mergeRegionParam.setRelaStartRow(craddr.getFirstRow() - curRowNum);
                                mergeRegionParam.setRelaEndRow(craddr.getLastRow() - curRowNum);
                                mergeRegionParam.setStartCol(craddr.getFirstColumn());
                                mergeRegionParam.setEndCol(craddr.getLastColumn());

                                allCellRangeAddress.put(idx, mergeRegionParam);
                            }
                        }
                        tagData.setAllCellRangeAddress(allCellRangeAddress);
                    }

                    curRowNum = curRowEndNum;
                    break;
                default:
                    throw new PoiExpandException("不支持的标签格式");
            }

            if(null == rootTagData){
                writeBlockMap.put(String.valueOf(writeBlockNum), tagData);
            } else {
                rootTagData.getChildTagDatas().add(tagData);
            }

            writeBlockNum ++;
            curRowNum ++;
        }
    }

    public static String getFirstCellValueStr(RowData rowData){
        String expr = null == rowData || MapUtil.isEmpty(rowData.getCellDatas()) || null == rowData.getCellDatas().get("0") ?
                null : String.valueOf(rowData.getCellDatas().get("0").getValue());
        return expr;
    };

    /**
     * 写入Excel Cell数据和样式
     *   foreach和bigforeach标签在写入样式时，使用了缓存的样式，避免创建很多相同的样式导致excel文件过大
     * @param writeWb
     * @param writeRow
     * @param readRowNum
     * @param cellData
     * @param params java动态参数
     * @param writeCellStyleCache 样式缓存
     */
    public static void writeCellData(Workbook writeWb, Row writeRow, int readRowNum, CellData cellData,
                                     Map<String, Object> params, Map<String, CellStyle> writeCellStyleCache){
        Cell writeCell = writeRow.createCell(cellData.getColNum());

        String cellStyleKey = readRowNum + "_" + cellData.getColNum();
        CellStyle cellStyle = null;
        if(null != writeCellStyleCache.get(cellStyleKey)){
            cellStyle = writeCellStyleCache.get(cellStyleKey);
            writeCell.setCellStyle(cellStyle);
            writeCell.setCellType(cellData.getCellType());
        } else {
            if(null != cellData.getCellStyle()){
                cellStyle = writeWb.createCellStyle();
                cellStyle.cloneStyleFrom(cellData.getCellStyle());
                writeCellStyleCache.put(cellStyleKey, cellStyle);
                writeCell.setCellStyle(cellStyle);
                writeCell.setCellType(cellData.getCellType());
            }
        }

        Object parseValue = ExprUtil.parseTempStr(params, cellData.getValue());
        ExcelCommonUtil.setCellValue(writeCell, parseValue);
    }

}
