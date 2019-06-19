package cn.sstool.poiexpand.common.bean.write;

import cn.sstool.poiexpand.common.bean.write.tag.TagData;
import lombok.Data;

import java.util.Map;

/**
 * 编  号：
 * 名  称：WriteSheetData
 * 描  述：
 * 完成日期：2019/6/19 23:28
 * @author：felix.shao
 */
@Data
public class WriteSheetData {

    /** 第几个sheet */
    private int sheetNum;

    /** sheet名称 */
    private String sheetName;

    /** 当前写入行号 */
    private int curWriteRowNum;

    /** 当前写入列号 */
    private int curWriteColNum;

    /** 列宽度 */
    private int[] cellWidths;

    /** 需要写入的数据块，块以TagData未单位 */
    private Map<String, TagData> writeTagDatas;

    /**
     * 根据读取的模板内容生成写入模板内容时使用。<br />
     * SXSSFSheet是一行行的写入数据，我们根据读取的模板生成需要写入的数据模板。
     * @return
     */
    public int getCurWriteRowNumAndIncrement(){
        return curWriteRowNum++;
    }

}
