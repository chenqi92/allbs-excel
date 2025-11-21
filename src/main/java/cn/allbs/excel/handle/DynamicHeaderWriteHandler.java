package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.DynamicHeaders;
import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import cn.allbs.excel.util.DynamicHeaderProcessor;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * DynamicHeader 处理器
 * <p>
 * 自动检测实体类中的 @DynamicHeaders 注解，并使用 DynamicHeaderProcessor 处理动态表头
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-21
 */
public class DynamicHeaderWriteHandler extends AbstractSheetWriteHandler {

    public DynamicHeaderWriteHandler(ExcelConfigProperties configProperties,
                                     ObjectProvider<List<Converter<?>>> converterProvider,
                                     WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 支持条件：
     * 1. 返回值是 List
     * 2. List 不为空
     * 3. List 元素不是 List
     * 4. List 元素的类中有字段标注了 @DynamicHeaders 注解
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

        // 检查是否有 @DynamicHeaders 注解
        Class<?> dataClass = objList.get(0).getClass();
        return hasDynamicHeadersAnnotation(dataClass);
    }

    @Override
    @SuppressWarnings("unchecked")
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> objList = (List<?>) obj;

        if (objList.isEmpty()) {
            throw new ExcelException("DynamicHeaders 导出不支持空数据");
        }

        Class<?> dataClass = objList.get(0).getClass();

        // 1. 展开数据
        List<Map<String, Object>> expandedData = DynamicHeaderProcessor.expandData((List<Object>) objList);

        // 2. 生成元数据（需要传入数据列表以提取动态表头）
        DynamicHeaderProcessor.DynamicHeaderMetadata metadata =
            DynamicHeaderProcessor.analyzeClass(dataClass, objList);

        // 3. 生成表头
        List<String> headers = DynamicHeaderProcessor.generateHeaders(metadata);
        List<List<String>> head = headers.stream()
            .map(Collections::singletonList)
            .collect(Collectors.toList());

        // 4. 转换数据为 List<List<Object>>
        List<List<Object>> rows = new ArrayList<>();
        for (Map<String, Object> rowMap : expandedData) {
            List<Object> row = new ArrayList<>();
            for (String header : headers) {
                row.add(rowMap.get(header));
            }
            rows.add(row);
        }

        // 获取 ExcelWriter
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        try {
            String sheetName = responseExcel.sheets().length > 0
                ? responseExcel.sheets()[0].sheetName()
                : "数据";

            // 创建 WriteSheet（使用自定义表头）
            WriteSheet writeSheet = EasyExcel.writerSheet(sheetName)
                .head(head)
                .build();

            // 写入数据
            excelWriter.write(rows, writeSheet);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 检查类中是否有字段标注了 @DynamicHeaders 注解
     *
     * @param clazz 数据类
     * @return true 有注解，false 无注解
     */
    private boolean hasDynamicHeadersAnnotation(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            DynamicHeaders dynamicHeaders = field.getAnnotation(DynamicHeaders.class);
            if (dynamicHeaders != null && dynamicHeaders.enabled()) {
                return true;
            }
        }
        return false;
    }
}
