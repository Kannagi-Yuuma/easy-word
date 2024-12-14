package com.yuuma.write;

import org.apache.poi.xwpf.usermodel.*;
import java.util.Map;
import java.util.List;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

/**
 * 表格内容写入器
 * 用于处理Word文档中的表格替换操作
 */
public class EasyWordTableWriter {

    /** 占位符正则表达式 */
    private String regex;

    /**
     * 创建EasyWordTableWriter实例
     * @param regex 占位符正则表达式 可选常用常量{link PlaceholderRegexConstant}
     */
    public EasyWordTableWriter(String regex) {
        this.regex = regex;
    }

    /**
     * 替换文档中表格的占位符
     * 
     * @param document Word文档对象
     * @param params   替换参数
     */
    public void replaceTablePlaceholders(XWPFDocument document, Map<String, Object> params) {
        // 获取所有表格
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            processTable(table, params);
        }
    }

    /**
     * 处理单个表格的占位符
     * 
     * @param table  表格对象
     * @param params 替换参数
     */
    private void processTable(XWPFTable table, Map<String, Object> params) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                // 获取单元格中的所有段落
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    // 获取段落中的所有运行块
                    List<XWPFRun> runs = paragraph.getRuns();
                    // 收集所有文本
                    StringBuilder text = new StringBuilder();
                    for (XWPFRun run : runs) {
                        text.append(run.getText(0));
                    }

                    // 替换占位符
                    String newText = replacePlaceholders(text.toString(), params);

                    // 如果文本发生变化，更新第一个运行块，删除其他运行块
                    if (!text.toString().equals(newText)) {
                        if (runs.size() > 0) {
                            XWPFRun firstRun = runs.get(0);
                            firstRun.setText(newText, 0);
                            // 删除其他运行块
                            for (int i = runs.size() - 1; i > 0; i--) {
                                paragraph.removeRun(i);
                            }
                        } else {
                            // 如果没有运行块，创建一个新的
                            XWPFRun run = paragraph.createRun();
                            run.setText(newText);
                        }
                    }
                }
            }
        }
    }

    /**
     * 替换文本中的占位符
     * 
     * @param text   原文本
     * @param params 替换参数
     * @return 替换后的文本
     */
    private String replacePlaceholders(String text, Map<String, Object> params) {
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        StringBuffer sb = new StringBuffer();

        while (matcher.find()) {
            String field = matcher.group(1);
            Object value = params.get(field);
            // 如果map中没有对应的值，使用空字符串
            String replacement = value != null ? value.toString() : "";
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(sb);

        return sb.toString();
    }

    /**
     * 设置正则表达式
     * @param regex 占位符正则表达式 可选常用常量{link PlaceholderRegexConstant}
     */
    public void setRegex(String regex) {
        this.regex = regex;
    }
}
