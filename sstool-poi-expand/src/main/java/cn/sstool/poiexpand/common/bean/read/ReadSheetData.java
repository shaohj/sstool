package cn.sstool.poiexpand.common.bean.read;

import lombok.Data;

import java.util.Map;

/**
 * 编  号：
 * 名  称：ReadSheetData
 * 描  述：
 * 完成日期：2019/6/19 23:26
 * @author：felix.shao
 */
@Data
public class ReadSheetData {

    /** 第几个sheet */
    private int sheetNum;

    /** sheet名称 */
    private String sheetName;

    /** 列宽度 */
    private int[] cellWidths;

    /** 读取sheet对应的行数据 */
    private Map<String, RowData> rowDatas;

}
