package com.github.shaohj.sstool.poiexpand.sax07;

import java.util.Arrays;
import java.util.stream.Collectors;

import static com.github.shaohj.sstool.poiexpand.common.consts.SaxExcelConst.FORMULA_KEY;
import static com.github.shaohj.sstool.poiexpand.common.consts.SaxExcelConst.FORMULA_THIS_ROW;

/**
 * 编  号：
 * 名  称：FormulaUtil
 * 描  述：
 * 完成日期：2019/12/14 18:09
 *
 * @author：felix.shao
 */
public class FormulaUtil {

    /**
     * 模板导出支持公式配置,如求和时,需要将空格等非数字字符转为数字才能继续处理
     * @param srcVal if条件需要替换的源字符 如-
     * @param tarVal if为true时需要替换的目标字符 如0
     * @param opt 操作符,如+ - * /
     * @param cellColumns 列编号,如A B C
     * @return
     */
    public static String cellOptByCurrentRow(String srcVal, String tarVal, String opt, String... cellColumns){
        return cellOptByRow(FORMULA_THIS_ROW, srcVal, tarVal, opt, cellColumns);
    }

    /**
     * 模板导出支持公式配置
     * @param row 行号
     * @param srcVal if条件需要替换的源字符串 如-
     * @param tarVal if为true时需要替换的目标字符串 如0
     * @param opt 操作符,如+ - * /
     * @param cellColumns 列编号,如A B C
     * @return
     */
    public static String cellOptByRow(String row, String srcVal, String tarVal, String opt, String... cellColumns){
        return cellOptByRowByDefault(row, srcVal, tarVal, opt, "", cellColumns);
    }

    /**
     * 模板导出支持公式配置
     * @param row 行号
     * @param srcVal if条件需要替换的源字符串 如-
     * @param tarVal if为true时需要替换的目标字符串 如0
     * @param opt 操作符,如+ - * /
     * @param nullDefault 参数不合理时的默认值
     * @param cellColumns 列编号,如A B C
     * @return
     */
    public static String cellOptByRowByDefault(String row, String srcVal, String tarVal, String opt, String nullDefault, String... cellColumns){
        if(null == cellColumns || cellColumns.length == 0){
            return nullDefault;
        }
        return  FORMULA_KEY + Arrays.asList(cellColumns).stream()
                .map(cellCol -> "IF(" + (cellCol + row) + "=\"" + srcVal + "\", " + tarVal + ", " + (cellCol + row) + ")")
                .collect(Collectors.joining(opt));
    }

}
