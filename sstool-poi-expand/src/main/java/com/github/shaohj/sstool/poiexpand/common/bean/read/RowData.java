package com.github.shaohj.sstool.poiexpand.common.bean.read;

import lombok.Data;

import java.util.Map;

/**
 * 编  号：
 * 名  称：RowData
 * 描  述：
 * 完成日期：2019/6/19 23:27
 * @author：felix.shao
 */
@Data
public class RowData {

    /** 行号 */
    private int rowNum;

    /** 读取行数据对应的列数据 */
    private Map<String, CellData> cellDatas;

    /** 对应Row.getHeight */
    private short height;

    /** 对应Row.getHeightInPoints */
    private float heightInPoints;

}
