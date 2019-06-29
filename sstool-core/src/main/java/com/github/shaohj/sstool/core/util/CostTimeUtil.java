package com.github.shaohj.sstool.core.util;

import com.github.shaohj.sstool.core.util.func.CostTimeFunc;
import lombok.extern.slf4j.Slf4j;

import java.time.Duration;
import java.time.Instant;

/**
 * 编  号：
 * 名  称：CostTimeUtil
 * 描  述：
 * 完成日期：2019/06/17 19:03
 *
 * @author：felix.shao
 */
@Slf4j
public class CostTimeUtil {

    public static <T> void apply(T t, String logTemp, CostTimeFunc<T> func){
        Instant start = Instant.now();
        func.apply(t);
        log.info(logTemp, Duration.between(start, Instant.now()).toMillis() + " ms");
    }

}
