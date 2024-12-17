package top.yuuma.entity;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import lombok.Data;
import lombok.NoArgsConstructor;
import top.yuuma.write.EasyWordWriter;
import lombok.AllArgsConstructor;

/**
 * Word文档实体类，作为XWPFDocument的包装类
 * 用于在框架中传递和操作Word文档数据
 *
 * @author yuuma
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class EasyWordDocument {

    /**
     * Apache POI的Word文档对象
     */
    private XWPFDocument document;

    /**
     * 静态工厂方法，用于创建EasyWordDocument实例
     *
     * @param document Apache POI的XWPFDocument对象
     * @return 返回包装后的EasyWordDocument对象
     */
    public static EasyWordDocument of(XWPFDocument document)  {
        EasyWordDocument easyWordDocument = new EasyWordDocument();
        easyWordDocument.setDocument(document);
        return easyWordDocument;
    }

    /**
     * 获取文档写入器
     * 用于进行文档内容的写入操作
     *
     * @return 返回与当前文档关联的EasyWordWriter实例
     */
    public EasyWordWriter write() {
        return EasyWordWriter.of(document);
    }
}
