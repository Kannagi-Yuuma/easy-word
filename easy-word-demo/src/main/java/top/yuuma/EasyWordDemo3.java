package top.yuuma;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

import lombok.extern.slf4j.Slf4j;
import top.yuuma.entity.EasyWordDocument;
import top.yuuma.read.EasyWordReader;

@Slf4j
public class EasyWordDemo3 {

    public static void main(String[] args) {
        // 占位符参数
        Map<String, Object> params = new HashMap<>();
        params.put("createTime", OffsetDateTime.now().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日")));
        params.put("arrivalTime", OffsetDateTime.now().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日")));
        params.put("name", "张三");
        params.put("phone", "13800138000");
        params.put("originalProject", "原项目名称");
        params.put("originalPost", "原岗位名称");
        params.put("originalAddress", "原地址");
        params.put("project", "新项目名称");
        params.put("post", "新岗位名称");
        params.put("address", "新地址");
        params.put("contactName", "联系人姓名");
        params.put("contactPhone", "18900189000");

        try (FileOutputStream out = new FileOutputStream("demo_template3_output.docx")) {
            // 读取Word文档
            EasyWordDocument document = EasyWordReader.read("file/demo_template3.docx");
            // 替换占位符数据,使用{link PlaceholderRegexConstant#DOUBLE_BRACE}正则表达式匹配占位符
            document.write()
                    .replacePlaceholders(params)
                    .export(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
