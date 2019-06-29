package com.github.shaohj.sstool.poiexpand.common.sax07;

import com.github.shaohj.sstool.core.util.CostTimeUtil;
import com.github.shaohj.sstool.poiexpand.common.entity.ModelTest;
import com.github.shaohj.sstool.poiexpand.sax07.Sax07ExcelExportParam;
import com.github.shaohj.sstool.poiexpand.sax07.Sax07ExcelUtil;
import com.github.shaohj.sstool.poiexpand.sax07.service.Sax07ExcelPageWriteService;
import lombok.extern.slf4j.Slf4j;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.*;

/**
 * 编  号：
 * 名  称：Sax07ExcelUtilTest
 * 描  述：
 * 完成日期：2019/01/29 00:27
 *
 * @author：felix.shao
 */
@Slf4j
public class Sax07ExcelUtilTest {

    public static final String exportPath = "E:\\temp\\export\\";

    @Before
    public void before(){
        log.info("缓存导出的xlsx临时文件目录为:{}", System.getProperty("java.io.tmpdir"));
    }

    /**
     * javabean为object的导出测试
     */
    @Test
    public void exportEachTest(){
        CostTimeUtil.apply(null, "导出1条数据,耗费时间为{}毫秒", s -> {
            String tempPath = "xlsx/";

            Map<String, Object> params = new HashMap<>();
            params.put("printDate", "2019-01-31");

            ModelTest model = new ModelTest("aaa1", "bbb", 123.234);
            model.setYear("1992");
            params.put("model", model);

            try(FileOutputStream fos = new FileOutputStream (exportPath + "each_data.xlsx")) {
                Sax07ExcelExportParam param = Sax07ExcelExportParam.builder()
                        .tempIsClassPath(true)
                        .tempFileName(tempPath + "each_temp.xlsx")
                        .params(params)
                        .outputStream(fos).build();

                Sax07ExcelUtil.export(param);
            }  catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    /**
     * javabean为map的导出测试
     */
    @Test
    public void exportEachMapTest(){
        CostTimeUtil.apply(null, "导出1条数据,耗费时间为{}毫秒", s -> {
            String tempPath = "xlsx/";

            Map<String, Object> params = new HashMap<>();
            params.put("printDate", "2019-01-31");

            Map model = new LinkedHashMap();
            model.put("uname", "zhangsan");
            model.put("realname", "张三");
            model.put("age", 19);

            params.put("model", model);

            try(FileOutputStream fos = new FileOutputStream (exportPath + "each_data.xlsx")) {
                Sax07ExcelExportParam param = Sax07ExcelExportParam.builder()
                        .tempIsClassPath(true)
                        .tempFileName(tempPath + "each_temp.xlsx")
                        .params(params)
                        .outputStream(fos).build();

                Sax07ExcelUtil.export(param);
            }  catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    @Test
    public void exportForeachTest(){
        int num = 4;
        CostTimeUtil.apply(null, "导出" + num + "条数据,耗费时间为{}毫秒", s -> {
            String tempPath = "xlsx/";

            Map<String, Object> params = new HashMap<>();
            params.put("printDate", "2019-01-31");

            ModelTest model = new ModelTest("aaa1", "bbb", 123.234);
            model.setYear("1992");
            params.put("model", model);

            List<ModelTest> details = new ArrayList<>();
            //i条数据导出测试
            for(int i = 0; i< num; i++){
                details.add(new ModelTest("user" + i, "world", 144.342));
            }
            params.put("list", details);

            try (FileOutputStream fos = new FileOutputStream(exportPath + "foreach_data.xlsx")){
                Sax07ExcelExportParam param = Sax07ExcelExportParam.builder()
                        .tempIsClassPath(true)
                        .tempFileName(tempPath + "foreach_temp.xlsx")
                        .params(params)
                        .outputStream(fos).build();
                Sax07ExcelUtil.export(param);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

    }

    @Test
    public void exportNestIfAndForeachTest(){
        int num = 4;
        CostTimeUtil.apply(null, "导出" + num + "条数据,耗费时间为{}毫秒", s -> {
            String tempPath = "xlsx/";

            Map<String, Object> params = new HashMap<>();
            params.put("printDate", "2019-01-31");

            ModelTest model = new ModelTest("aaa1", "bbb", 123.234);
            model.setYear("1992");
            params.put("model", model);

            List details = new ArrayList();
            //i条数据导出测试
            for(int i = 0; i< num; i++){
                details.add(new ModelTest("user" + i, "world", 144.342));
            }
            params.put("list", details);

            try (FileOutputStream fos = new FileOutputStream(exportPath + "nest_if_foreach_data.xlsx")){
                Sax07ExcelExportParam param = Sax07ExcelExportParam.builder()
                        .tempIsClassPath(true)
                        .tempFileName(tempPath + "nest_if_foreach_temp.xlsx")
                        .params(params)
                        .outputStream(fos).build();
                Sax07ExcelUtil.export(param);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    @Test
    public void exportPageForeachTest(){
        int num = 4;
        CostTimeUtil.apply(null, "导出" + num + "条数据,耗费时间为{}毫秒", s -> {
            int pageSize = 10;
            int totalPageNum = num % pageSize == 0 ? num/pageSize : num/pageSize + 1;

            String tempPath = "xlsx/";

            Map<String, Object> params = new HashMap<>();
            params.put("printDate", "2019-01-31");

            ModelTest model = new ModelTest("aaa1", "bbb", 123.234);
            model.setYear("1992");
            params.put("model", model);

            Sax07ExcelPageWriteService sax07ExcelPageWriteService = new Sax07ExcelPageWriteService(){
                @Override
                public void pageWriteData() {
                    for (int i = 0; i <totalPageNum; i++) {
                        Map<String, Object> pageParams = new HashMap<>();
                        List details = new ArrayList(pageSize);
                        for (int j = 0; j <pageSize && pageSize * i + j < num; j++) {
                            details.add(new ModelTest("user" + j, "world", 144.342));
                        }
                        pageParams.put("list", details);
                        tagData.writeTagData(writeWb, writeSheet, writeSheetData, pageParams, writeCellStyleCache);
                    }
                }
            };
            //设置sax07ExcelPageWriteService对应的表达式
            sax07ExcelPageWriteService.setExprVal("#pageforeach detail in ${list}");

            try (FileOutputStream fos = new FileOutputStream(exportPath + "pageforeach_data.xlsx");){
                Sax07ExcelExportParam param = Sax07ExcelExportParam.builder()
                        .tempIsClassPath(true)
                        .tempFileName(tempPath + "pageforeach_temp.xlsx")
                        .params(params)
                        .sax07ExcelPageWriteServices(Arrays.asList(sax07ExcelPageWriteService))
                        .outputStream(fos).build();

                Sax07ExcelUtil.export(param);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

}
