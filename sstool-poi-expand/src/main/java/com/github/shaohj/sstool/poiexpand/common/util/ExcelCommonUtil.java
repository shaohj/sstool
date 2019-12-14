package com.github.shaohj.sstool.poiexpand.common.util;

import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.core.util.StrUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.write.MergeRegionParam;
import com.github.shaohj.sstool.poiexpand.common.consts.SaxExcelConst;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMergeCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMergeCells;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static com.github.shaohj.sstool.poiexpand.common.consts.SaxExcelConst.FORMULA_KEY;

/**
 * 编  号：
 * 名  称：ExcelCommonUtil
 * 描  述：
 * 完成日期：2019/6/19 23:53
 * @author：felix.shao
 */
@Slf4j
public class ExcelCommonUtil {

    /**
     * 获取Sheet最大的数据列数
     * @param sheet :
     * @return int  最大列数为3则返回3
     * @author SHJ
     */
    public static int getMaxCellNum(Sheet sheet){
        int rows = sheet.getPhysicalNumberOfRows();
        int maxCellNum = 0;

        for(int i=0; i< rows; i++){
            Row tRow = sheet.getRow(i);
            if(null != tRow){
                int tempNum = tRow.getLastCellNum();
                maxCellNum = tempNum > maxCellNum ? tempNum : maxCellNum;
            }
        }
        return maxCellNum;
    }

    /**
     * 获取excel所有的列宽
     * @param sheet
     * @param maxCellNum
     * @return
     */
    public static int[] getCellWidths(Sheet sheet, int maxCellNum){
        int[] cellWidths = new int[maxCellNum];

        for (int i = 0; i < maxCellNum; i++) {
            cellWidths[i] = sheet.getColumnWidth(i);
        }

        return cellWidths;
    }

    public static Object getCellValue(Cell cell){
        Object result = null;
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                result = cell.getBooleanCellValue(); break;
            case FORMULA:
                result = FORMULA_KEY + cell.getCellFormula(); break;
            case NUMERIC:
                result = cell.getNumericCellValue(); break;
            case STRING:
                result = cell.getStringCellValue(); break;
            default:
        }
        return result;
    }

    public static void setCellValue(Cell writeCell, Object value, int currentWriteRow){
        if(null == value){
            writeCell.setCellValue("");
        } else {
            if ("java.lang.Integer".equals(value.getClass().getName())) {
                writeCell.setCellValue(Double.parseDouble(value.toString()));
            } else if ("java.lang.Double".equals(value.getClass().getName())) {
                writeCell.setCellValue(((Double) value).doubleValue());
            } else if ("java.util.Date".equals(value.getClass().getName())) {
                writeCell.setCellValue((Date) value);
            } else if ("java.lang.Boolean".equals(value.getClass().getName())) {
                writeCell.setCellValue(((Boolean) value).booleanValue());
            } else if ("java.lang.String".equals(value.getClass().getName())) {
                int formulaIndex = value.toString().indexOf(FORMULA_KEY);
                //  formulaIndex为0时,表示string为读取后处理后的公式数据或自定义的公式数据
                if(0 != formulaIndex){
                    writeCell.setCellValue(value.toString());
                } else {
                    String formulaValue = value.toString();
                    formulaValue = formulaValue.substring(FORMULA_KEY.length());
                    // 当前写入行号,后续考虑兼容基于该行号相对配置其他行号功能
                    int actualCurrentWriteRow = currentWriteRow + 1;
                    formulaValue = formulaValue.replaceAll(SaxExcelConst.FORMULA_THIS_ROW, String.valueOf(actualCurrentWriteRow));
                    // String如果是公式,则设置公式值
                    writeCell.setCellFormula(formulaValue);
                }
            } else {
                writeCell.setCellValue(value.toString());
            }
        }
    }

    /**
     * 定制化的合并单元格优化
     * 对比原有合并单元格，去掉了验证逻辑和返回值
     * @param sxssfWorkbok
     * @param sheetName
     * @param region
     * @param mergeCellsCount
     */
    public static void addMergeRegion(SXSSFWorkbook sxssfWorkbok, String sheetName, CellRangeAddress region, int mergeCellsCount) {
        XSSFSheet sheet = sxssfWorkbok.getXSSFWorkbook().getSheet(sheetName);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();

        CTMergeCells ctMergeCells = mergeCellsCount > 0 ?ctWorksheet.getMergeCells():ctWorksheet.addNewMergeCells();
        CTMergeCell ctMergeCell = ctMergeCells.addNewMergeCell();

        ctMergeCell.setRef(region.formatAsString());
    }

    /**
     * 获取动态设置的合并单元格参数
     * @param rowData eq:#each ${model} mergeRegion=0,1,0,2|0,1,2,4  即第1行第1和2列合并，第1行第3和4列合并，多个用|隔开
     * @return MergeRegionParam
     */
    public static Map<String, MergeRegionParam> getMergeRegionParam(String rowData){
        if(StrUtil.isEmpty(rowData)){
            return new HashMap<>(0);
        }
        int idx = rowData.lastIndexOf(SaxExcelConst.MERGE_REGION_OPT);
        if(idx < 0){
            return new HashMap<>(0);
        }
        String params = rowData.substring(idx + SaxExcelConst.MERGE_REGION_OPT.length());
        if(StrUtil.isEmpty(params)){
            return new HashMap<>(0);
        }
        String[] mRegionParams = params.split("[|]");

        Map<String, MergeRegionParam> result = new HashMap<>(MapUtil.calMapSize(mRegionParams.length));

        for (int i = 0; i < mRegionParams.length; i++ ){
            String[] mrTemp = mRegionParams[i].split(",");
            if(4 != mrTemp.length){
                log.warn("mergeRegion param config error, ignore the str = {}", mRegionParams[i]);
            } else {
                MergeRegionParam mrParam = new MergeRegionParam();
                mrParam.setRelaStartRow(Integer.parseInt(mrTemp[0]));
                mrParam.setRelaEndRow(Integer.parseInt(mrTemp[1]));
                mrParam.setStartCol(Integer.parseInt(mrTemp[2]));
                mrParam.setEndCol(Integer.parseInt(mrTemp[3]));

                result.put("dy_" + i , mrParam);
            }
        }

        return result;
    }

}
