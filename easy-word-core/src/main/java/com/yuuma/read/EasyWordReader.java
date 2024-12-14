package com.yuuma.read;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileNotFoundException;

import com.yuuma.entity.EasyWordDocument;

/**
 * Word文档读取工具类
 * 用于读取Word文档并返回EasyWordDocument对象
 */
public class EasyWordReader {

    /**
     * 读取Word文档
     * 
     * @param fileName 文件名 (在resources目录下的文件名)
     * @return XWPFDocument对象
     * @throws IOException 如果文件读取失败
     */
    public static EasyWordDocument read(String fileName) throws IOException {
        // 获取资源文件
        ClassLoader classLoader = EasyWordReader.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream(fileName);

        if (inputStream == null) {
            throw new FileNotFoundException("文件未在resources目录下找到: " + fileName);
        }

        // 根据文件扩展名判断文件名后缀是否合规
        if (fileName.toLowerCase().endsWith(".docx")) {
            // 处理.docx文件
            return EasyWordDocument.of(new XWPFDocument(inputStream));
        } else {
            throw new IllegalArgumentException("不支持的文件格式，仅支持.docx文件");
        }
    }

    /**
     * 读取Word文档
     * 
     * @param inputStream 输入流
     * @return EasyWordDocument对象
     * @throws IOException 如果文件读取失败
     */
    public static EasyWordDocument read(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new FileNotFoundException("输入流为空");
        }
        return EasyWordDocument.of(new XWPFDocument(inputStream));
    }
}
