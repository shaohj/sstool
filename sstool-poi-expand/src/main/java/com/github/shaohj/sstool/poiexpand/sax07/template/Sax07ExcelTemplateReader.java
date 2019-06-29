package com.github.shaohj.sstool.poiexpand.sax07.template;

import com.github.shaohj.sstool.core.util.IoUtil;
import com.github.shaohj.sstool.core.util.MapUtil;
import com.github.shaohj.sstool.core.util.StrUtil;
import com.github.shaohj.sstool.poiexpand.common.bean.read.CellData;
import com.github.shaohj.sstool.poiexpand.common.bean.read.ReadSheetData;
import com.github.shaohj.sstool.poiexpand.common.bean.read.RowData;
import com.github.shaohj.sstool.poiexpand.common.exception.PoiExpandException;
import com.github.shaohj.sstool.poiexpand.common.util.ExcelCommonUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 编  号：
 * 名  称：Sax07ExcelTemplateReader
 * 描  述：excle2007读取模板工具类
 * 完成日期：2019/6/29 11:11
 * @author：felix.shao
 */
public class Sax07ExcelTemplateReader {

    /**
     * 根据模板文件名打开workbook.
     * @param isClassPath 是否为class路径.true:是;false:否
     * @param fileName 文件路径，支持class路径和绝对路径
     * @return
     */
    public static XSSFWorkbook openWorkbook(boolean isClassPath, String fileName) {
        String fileSuffix = StrUtil.getStringSuffix(fileName);
        if (!"xlsx".equals(fileSuffix)) {
            throw new IllegalArgumentException("poi缓存导出只支持xlsx,程序同步,支持支xlsx导出");
        }
        InputStream in = null;
        try {
            if(isClassPath) {
                in = Sax07ExcelTemplateReader.class.getClassLoader().getResourceAsStream(fileName);
            } else {
                in = new FileInputStream(fileName);
            }
            return new XSSFWorkbook(in);
        } catch (Exception e) {
            throw new PoiExpandException("File" + fileName + "not found" + e.getMessage());
        } finally {
            IoUtil.close(in);
        }
    }

    /**
     * 读取XSSFWorkbook的Sheet，生成SheetData数据
     * @param readWb
     * @return
     */
    public static List<ReadSheetData> readSheetData(Workbook readWb){
        //模板中所有sheet数量
        int readWbSheetCount = readWb.getNumberOfSheets();
        List<ReadSheetData> readSheetDatas = new ArrayList<>(readWbSheetCount);

        for (int i = 0; i < readWbSheetCount; i++) {
            Sheet readSheet = readWb.getSheetAt(i);

            ReadSheetData readSheetData = new ReadSheetData();
            readSheetData.setSheetNum(i);
            readSheetData.setSheetName(readSheet.getSheetName());

            int maxCellNum = ExcelCommonUtil.getMaxCellNum(readSheet);
            int[] cellWidths = ExcelCommonUtil.getCellWidths(readSheet, maxCellNum);
            readSheetData.setCellWidths(cellWidths);

            Map<String, RowData> readSheetRowDatas = readRowData(readSheet, readSheet.getFirstRowNum(), readSheet.getLastRowNum());
            readSheetData.setRowDatas(readSheetRowDatas);

            readSheetDatas.add(readSheetData);
        }

        return readSheetDatas;
    }

    /**
     * 读取XSSFWorkbook的Row，生成RowData数据
     * @param readSheet
     * @param firstRowNum
     * @param lastRowNum
     * @return
     */
    public static Map<String, RowData> readRowData(Sheet readSheet, int firstRowNum, int lastRowNum){
        Map<String, RowData> rowDatas = new LinkedHashMap<>(MapUtil.calMapCapacity(lastRowNum - firstRowNum + 1));
        int curRowNum = firstRowNum;

        while (curRowNum <= lastRowNum) {
            Row readRow = readSheet.getRow(curRowNum);

            RowData rowData = new RowData();
            rowData.setRowNum(curRowNum);

            if(null != readRow){
                //设置行高等属性
                rowData.setHeight(readRow.getHeight());
                rowData.setHeightInPoints(readRow.getHeightInPoints());
            }

            Map<String, CellData> readRowCellDatas = null != readRow ?
                    readCellData(readRow, readRow.getFirstCellNum(), readRow.getLastCellNum())
                    : new LinkedHashMap<>(0);

            rowData.setCellDatas(readRowCellDatas);

            rowDatas.put(String.valueOf(curRowNum), rowData);

            curRowNum++;
        }

        return rowDatas;
    }

    /**
     * 读取XSSFWorkbook的Cell，生成CellData数据
     * @param readRow
     * @param firstCellNum
     * @param lastCellNum
     * @return
     */
    public static Map<String, CellData> readCellData(Row readRow, int firstCellNum, int lastCellNum){
        Map<String, CellData> cellDatas = new LinkedHashMap<>(MapUtil.calMapCapacity(lastCellNum - firstCellNum + 1));

        if(-1 == firstCellNum) {
            return cellDatas;
        }
        int curCellNum = firstCellNum;

        while (curCellNum <= lastCellNum) {
            Cell readCell = readRow.getCell(curCellNum);

            CellData cellData = new CellData();
            cellData.setColNum(curCellNum);

            if(null != readCell){
                //设置样式等属性
                cellData.setCellStyle(readCell.getCellStyle());
                cellData.setCellType(readCell.getCellTypeEnum());
            }

            if (null == readCell){
                cellData.setValue("");
            } else {
                cellData.setValue(ExcelCommonUtil.getCellValue(readCell));
            }

            cellDatas.put(String.valueOf(curCellNum), cellData);

            curCellNum++;
        }

        return cellDatas;
    }

}
