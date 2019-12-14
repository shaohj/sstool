package com.github.shaohj.sstool.poiexpand.common.sax07;

import com.github.shaohj.sstool.poiexpand.sax07.FormulaUtil;
import org.junit.Test;

/**
 * 编  号：
 * 名  称：FormulaUtilTest
 * 描  述：
 * 完成日期：2019/12/14 18:37
 *
 * @author：felix.shao
 */
public class FormulaUtilTest {

    @Test
    public void cellOptByRowTest(){
        String row = "2";
        System.out.println(FormulaUtil.cellOptByRowByDefault(
                row, "-", "0", "+", "", "A", "B"));
    }

}
