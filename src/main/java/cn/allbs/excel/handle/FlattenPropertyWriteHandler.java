package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.annotation.FlattenProperty;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import cn.allbs.excel.util.FlattenFieldProcessor;
import cn.idev.excel.FastExcel;
import cn.idev.excel.ExcelWriter;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.write.metadata.WriteSheet;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * FlattenProperty 处理器
 * <p>
 * 自动检测实体类中的 @FlattenProperty 注解，并使用 FlattenFieldProcessor 处理嵌套对象展开
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-21
 */
public class FlattenPropertyWriteHandler extends AbstractSheetWriteHandler {

    public FlattenPropertyWriteHandler(ExcelConfigProperties configProperties,
                                       ObjectProvider<List<Converter<?>>> converterProvider,
                                       WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 支持条件：
     * 1. 返回值是 List
     * 2. List 不为空
     * 3. List 元素不是 List
     * 4. List 元素的类中有字段标注了 @FlattenProperty 注解
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

        // 检查是否有 @FlattenProperty 注解
        Class<?> dataClass = objList.get(0).getClass();
        return hasFlattenPropertyAnnotation(dataClass);
    }

    @Override
    @SuppressWarnings("unchecked")
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> objList = (List<?>) obj;

        if (objList.isEmpty()) {
            throw new ExcelException("FlattenProperty 导出不支持空数据");
        }

        Class<?> dataClass = objList.get(0).getClass();

        // 使用 FlattenFieldProcessor 分析类结构
        List<FlattenFieldProcessor.FlattenFieldInfo> fieldInfos = FlattenFieldProcessor.processFlattenFields(dataClass);

        // 生成表头
        List<List<String>> head = fieldInfos.stream()
            .map(info -> Collections.singletonList(info.getHeadName()))
            .collect(Collectors.toList());

        // 转换数据为行
        List<List<Object>> rows = new ArrayList<>();
        for (Object data : objList) {
            List<Object> row = new ArrayList<>();
            for (FlattenFieldProcessor.FlattenFieldInfo fieldInfo : fieldInfos) {
                Object value = FlattenFieldProcessor.extractValue(data, fieldInfo);
                row.add(value);
            }
            rows.add(row);
        }

        // 获取 ExcelWriter
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        try {
            // 创建 WriteSheet（使用自定义表头）
            WriteSheet writeSheet = FastExcel.writerSheet(
                responseExcel.sheets()[0].sheetName()
            ).head(head).build();

            // 写入数据
            excelWriter.write(rows, writeSheet);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 检查类中是否有字段标注了 @FlattenProperty 注解
     *
     * @param clazz 数据类
     * @return true 有注解，false 无注解
     */
    private boolean hasFlattenPropertyAnnotation(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            FlattenProperty flattenProperty = field.getAnnotation(FlattenProperty.class);
            if (flattenProperty != null) {
                return true;
            }
        }
        return false;
    }
}
