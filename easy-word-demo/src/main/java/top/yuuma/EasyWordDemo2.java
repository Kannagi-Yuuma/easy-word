package top.yuuma;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import lombok.extern.slf4j.Slf4j;
import top.yuuma.constants.PlaceholderRegexConstant;
import top.yuuma.entity.EasyWordDocument;
import top.yuuma.read.EasyWordReader;

@Slf4j
public class EasyWordDemo2 {

    public static void main(String[] args) {
        // 占位符参数
        Map<String, Object> params = new HashMap<>();
        params.put("time", new Date());
        params.put("text", "从早上开始上班，上到晚上11点");
        params.put("tableText1", "我是插进来的文本");
        params.put("tableText2", "我是有颜色的文本");

        try (FileOutputStream out = new FileOutputStream("demo_template2_output.docx")) {
            // 读取Word文档
            EasyWordDocument document = EasyWordReader.read("file/demo_template2.docx");
            // 替换占位符数据,使用{link PlaceholderRegexConstant#DOUBLE_BRACE}正则表达式匹配占位符
            document.write()
                    .replacePlaceholders(params, PlaceholderRegexConstant.DOUBLE_BRACE)
                    .export(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
