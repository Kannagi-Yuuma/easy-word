package top.yuuma;

import lombok.extern.slf4j.Slf4j;
import top.yuuma.entity.EasyWordDocument;
import top.yuuma.read.EasyWordReader;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * 示例代码
 */
@Slf4j
public class EasyWordDemo {
    public static void main(String[] args) {
        // 占位符参数
        Map<String, Object> params = new HashMap<>();
        params.put("time", new Date());
        params.put("text", "从早上开始上班，上到晚上11点");
        params.put("tableText1", "我是插进来的文本");
        params.put("tableText2", "我是有颜色的文本");

        try (FileOutputStream out = new FileOutputStream("demo_template_output.docx")) {
            // 读取Word文档
            EasyWordDocument document = EasyWordReader.read("file/demo_template.docx");
            // 替换占位符数据
            document.write()
                    .replacePlaceholders(params)
                    .export(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
