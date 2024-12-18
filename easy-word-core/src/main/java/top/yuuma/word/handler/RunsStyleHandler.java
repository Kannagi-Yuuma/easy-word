package top.yuuma.word.handler;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Run样式处理器
 * 用于处理Word文档中run的样式保存和恢复
 */
public class RunsStyleHandler {
    
    /**
     * 保存原有格式并设置新文本
     * @param run XWPFRun对象
     * @param newText 新文本
     */
    public static void preserveAndSetText(XWPFRun run, String newText) {
        // 保存文本格式
        boolean isBold = run.isBold();
        boolean isItalic = run.isItalic();
        boolean hasUnderline = run.getUnderline() != UnderlinePatterns.NONE;
        boolean isStrike = run.isStrikeThrough();
        UnderlinePatterns underlinePattern = run.getUnderline();
        
        // 字体相关
        String fontFamily = run.getFontFamily();
        Double fontSize = run.getFontSizeAsDouble();
        String color = run.getColor();
        
        // 设置新文本
        run.setText(newText, 0);
        
        // 恢复格式
        run.setBold(isBold);
        run.setItalic(isItalic);
        if (hasUnderline) {
            run.setUnderline(underlinePattern);
        }
        run.setStrikeThrough(isStrike);
        
        // 恢复字体相关
        if (fontFamily != null) {
            run.setFontFamily(fontFamily);
        }
        if (fontSize != null && fontSize != -1) {
            run.setFontSize(fontSize);
        }
        if (color != null) {
            run.setColor(color);
        }
    }
}
