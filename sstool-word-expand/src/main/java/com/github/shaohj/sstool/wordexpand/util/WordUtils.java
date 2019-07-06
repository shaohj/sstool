package com.github.shaohj.sstool.wordexpand.util;

import com.alibaba.fastjson.util.IOUtils;
import com.github.shaohj.sstool.core.util.FileUtils;
import com.github.shaohj.sstool.core.util.IoUtil;
import com.github.shaohj.sstool.core.util.Java8DateUtils;
import com.github.shaohj.sstool.wordexpand.freemarker.FreemarkerUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipFile;
import org.apache.tools.zip.ZipOutputStream;

import java.io.*;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Enumeration;
import java.util.zip.ZipException;

/**
 * 编  号：
 * 名  称：WordUtils
 * 描  述：
 * 完成日期：2018/11/18 22:39
 * @author：felix.shao
 */
@Slf4j
public class WordUtils {

    /**
     * 替换模板文件的word/document.xml,生成新的docx,String导出会长度溢出
     * @param docxExportParam 导出参数
     * @throws ZipException
     * @throws IOException
     */
    public static void parseDocx(DocxExportParam docxExportParam) throws ZipException, IOException {
        String fromPath = docxExportParam.getRootDocxTempFileFolderUrl() + docxExportParam.getTempFileUrl();
        URL url = new File(fromPath).toURI().toURL();

        ZipFile zipFile = null;
        ZipOutputStream zipout = null;

        /** 创建导出父文件夹 */
        FileUtils.createParentFolder(docxExportParam.getToFilePath());

        try {
            zipFile = new ZipFile(URLDecoder.decode(url.getPath(), "UTF-8"), "UTF-8");
            Enumeration<? extends ZipEntry> zipEntrys = zipFile.getEntries();
            zipout = new ZipOutputStream(new FileOutputStream(docxExportParam.getToFilePath()));
            zipout.setEncoding("UTF-8");
            int len = -1;
            byte[] buffer=new byte[1024];
            while(zipEntrys.hasMoreElements()) {
                ZipEntry next = zipEntrys.nextElement();
                InputStream is = zipFile.getInputStream(next);
                //把输入流的文件传到输出流中 如果是word/document.xml由我们输入
                zipout.putNextEntry(new ZipEntry(next.toString()));
                if("word/document.xml".equals(next.toString())){
                    /** 在模板文件目录下生成模板文件 */
                    String fileName = Java8DateUtils.getYyyyMmDDHhMmSs() + "_01.ftl";
                    String newFileName = Java8DateUtils.getYyyyMmDDHhMmSs() + "_01_new.ftl";
                    String ftlTempFile = docxExportParam.getRootFreemarkerTempFileFolderUrl() + "\\" + fileName;
                    String ftlNewFile = docxExportParam.getRootFreemarkerTempFileFolderUrl() + "\\" + newFileName;
                    FileUtils.createFile(ftlTempFile);
                    FileUtils.createFile(ftlNewFile);
                    File ftlFile = new File(ftlTempFile);
                    FileOutputStream ftlOut = null;
                    try {
                        ftlOut = new FileOutputStream(ftlFile);
                        IoUtil.copy(is, ftlOut);
                        IOUtils.close(ftlOut);
                        /** 替换模板表达式，生成替换后的文件 */
                        FreemarkerUtil.parseTFToFile(docxExportParam, fileName, ftlNewFile);
                        /** 将替换后的文件写入压缩文件中 */
                        InputStream in = new FileInputStream(ftlNewFile);
                        while((len = in.read(buffer))!=-1){
                            zipout.write(buffer,0,len);
                        }
                        IOUtils.close(in);
                    } finally {
                        IOUtils.close(ftlOut);
                        /** 删除临时文件 */
                        FileUtils.deleteFile(ftlTempFile);
                        FileUtils.deleteFile(ftlNewFile);
                    }
                }else {
                    while((len = is.read(buffer))!=-1){
                        zipout.write(buffer,0,len);
                    }
                    IOUtils.close(is);
                }
            }
            log.info("parseDocx success.\nfromPath-->{}\ntoPath-->[{}] success,", fromPath, docxExportParam.getToFilePath());
        } catch (FileNotFoundException e) {
            log.error("Error by parseDocx.{}", e);
        } finally{
            IOUtils.close(zipout);
            IOUtils.close(zipFile);
        }
    }

}
