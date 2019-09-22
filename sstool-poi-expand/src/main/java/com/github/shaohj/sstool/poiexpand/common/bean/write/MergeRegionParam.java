package com.github.shaohj.sstool.poiexpand.common.bean.write;

import lombok.Data;

/**
 * 编  号：
 * 名  称：MergeRegionParam
 * 描  述：
 * 完成日期：2019/09/20 16:45
 *
 * @author：felix.shao
 */
@Data
public class MergeRegionParam {

    /** 相对的起始行数 */
    private int relaStartRow;

    /** 相对的结束行数 */
    private int relaEndRow;

    /** 起始列数 */
    private int startCol;

    /** 结束列数 */
    private int endCol;

}
