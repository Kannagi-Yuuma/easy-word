package top.yuuma.word.write;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import top.yuuma.word.constants.PlaceholderRegexConstant;

import java.util.Map;
import java.io.OutputStream;
import java.io.IOException;
import java.io.ByteArrayOutputStream;

/**
 * Word文档写入工具类
 * 用于处理Word文档中的占位符替换和导出操作
 */
public class EasyWordWriter {

    /** 文本内容写入器 */
    private EasyWordTextWriter textWriter;
    /** 表格内容写入器 */
    private EasyWordTableWriter tableWriter;
    /** Word文档对象 */
    private XWPFDocument document;

    /**
     * 创建EasyWordWriter实例
     * @param document Word文档对象
     * @return EasyWordWriter实例
     */
    public static EasyWordWriter of(XWPFDocument document) {
        EasyWordWriter writer = new EasyWordWriter();
        writer.document = document;
        writer.textWriter = new EasyWordTextWriter(PlaceholderRegexConstant.NORMAL);
        writer.tableWriter = new EasyWordTableWriter(PlaceholderRegexConstant.NORMAL);
        return writer;
    }


    /**
     * 替换文档中的占位符,默认使用{link PlaceholderRegexConstant#NORMAL}正则表达式匹配占位符
     * @param params 参数映射，key为占位符，value为替换值
     * @return 当前实例
     */
    public EasyWordWriter replacePlaceholders(Map<String, Object> params) {
        textWriter.replaceTextPlaceholders(document, params);
        tableWriter.replaceTablePlaceholders(document, params);
        return this;
    }

    /**
     * 替换文档中的占位符
     * @param params 参数映射，key为占位符，value为替换值
     * @param regex 占位符正则表达式 可选常用常量{link PlaceholderRegexConstant}
     * @return 当前实例
     */
    public EasyWordWriter replacePlaceholders(Map<String, Object> params, String regex) {
        textWriter.setRegex(regex);
        tableWriter.setRegex(regex);
        textWriter.replaceTextPlaceholders(document, params);
        tableWriter.replaceTablePlaceholders(document, params);
        return this;
    }

    /**
     * 将处理后的文档导出到输出流
     * @param outputStream 输出流
     * @throws IOException IO异常
     */
    public void export(OutputStream outputStream) throws IOException {
        document.write(outputStream);
    }

    /**
     * 将处理后的文档导出到字节数组输出流
     * @return 包含文档内容的字节数组输出流
     * @throws IOException IO异常
     */
    public ByteArrayOutputStream exportToStream() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.write(outputStream);
        return outputStream;
    }
}

