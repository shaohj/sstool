package com.github.shaohj.sstool.core.util.func;

/**
 * 编  号：
 * 名  称：CostTimeFunc
 * 描  述：耗时计算的一个参数封装
 * 完成日期：2019/06/17 19:00
 *
 * @author：felix.shao
 */
@FunctionalInterface
public interface CostTimeFunc<T> {

    void apply(T t);

}
