package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.annotation.RelatedSheet;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import cn.allbs.excel.util.MultiSheetRelationProcessor;
import cn.idev.excel.ExcelWriter;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.write.handler.WorkbookWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.List;

/**
 * 关联 Sheet 导出处理器
 * <p>
 * 自动检测实体类中的 @RelatedSheet 注解，并使用 MultiSheetRelationProcessor 处理关联关系
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-21
 */
@Slf4j
public class RelatedSheetWriteHandler extends AbstractSheetWriteHandler {

    public RelatedSheetWriteHandler(ExcelConfigProperties configProperties,
                                    ObjectProvider<List<Converter<?>>> converterProvider,
                                    WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 支持条件：
     * 1. 返回值是 List
     * 2. List 不为空
     * 3. List 元素不是 List（区别于 ManySheetWriteHandler）
     * 4. List 元素的类中有字段标注了 @RelatedSheet 注解
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (!(obj instanceof List)) {
            throw new ExcelException("@ResponseExcel 返回值必须为List类型");
        }

        List<?> objList = (List<?>) obj;

        // 空 List 或元素是 List，交给其他 Handler 处理
        if (objList.isEmpty() || objList.get(0) instanceof List) {
            return false;
        }

        // 检查是否有 @RelatedSheet 注解
        Class<?> dataClass = objList.get(0).getClass();
        return hasRelatedSheetAnnotation(dataClass);
    }

    @Override
    @SuppressWarnings("unchecked")
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> objList = (List<?>) obj;

        if (objList.isEmpty()) {
            throw new ExcelException("关联 Sheet 导出不支持空数据");
        }

        // 获取数据类型
        Object firstElement = objList.get(0);
        @SuppressWarnings("unchecked")
        Class<Object> dataClass = (Class<Object>) firstElement.getClass();
        @SuppressWarnings("unchecked")
        List<Object> typedList = (List<Object>) objList;

        // 获取主 Sheet 名称
        String mainSheetName = responseExcel.sheets().length > 0
            ? responseExcel.sheets()[0].sheetName()
            : "数据";

        // 创建 ExcelWriter
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        // 存储超链接信息的容器
        final List<MultiSheetRelationProcessor.HyperlinkInfo>[] hyperlinkContainer = new List[1];

        // 创建 WorkbookWriteHandler 来在完成时应用超链接
        WorkbookWriteHandler hyperlinkHandler = new WorkbookWriteHandler() {
            @Override
            public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
                if (hyperlinkContainer[0] != null && !hyperlinkContainer[0].isEmpty()) {
                    Workbook workbook = writeWorkbookHolder.getWorkbook();
                    log.info("Applying {} hyperlinks to workbook", hyperlinkContainer[0].size());
                    MultiSheetRelationProcessor.applyHyperlinks(workbook, hyperlinkContainer[0]);
                } else {
                    log.debug("No hyperlinks to apply");
                }
            }
        };

        try {
            // 导出数据并获取超链接信息
            List<MultiSheetRelationProcessor.HyperlinkInfo> hyperlinks =
                MultiSheetRelationProcessor.exportWithRelations(
                    excelWriter,
                    typedList,
                    mainSheetName,
                    dataClass
                );

            // 保存超链接信息供后续应用
            hyperlinkContainer[0] = hyperlinks;
            log.debug("Exported {} sheets with {} hyperlinks",
                     responseExcel.sheets().length + 1, hyperlinks.size());

            // 通过反射访问 writeContext 并注册处理器
            // 因为 EasyExcel 的 API 限制，我们需要直接操作 workbook
            if (!hyperlinks.isEmpty()) {
                // 在 finish() 前应用超链接
                // 获取 workbook 并应用超链接
                Workbook workbook = excelWriter.writeContext().writeWorkbookHolder().getWorkbook();
                log.info("Applying {} hyperlinks before finish", hyperlinks.size());
                MultiSheetRelationProcessor.applyHyperlinks(workbook, hyperlinks);
            }

        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 检查类中是否有字段标注了 @RelatedSheet 注解
     *
     * @param clazz 数据类
     * @return true 有注解，false 无注解
     */
    private boolean hasRelatedSheetAnnotation(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            RelatedSheet relatedSheet = field.getAnnotation(RelatedSheet.class);
            if (relatedSheet != null && relatedSheet.enabled()) {
                return true;
            }
        }
        return false;
    }
}
