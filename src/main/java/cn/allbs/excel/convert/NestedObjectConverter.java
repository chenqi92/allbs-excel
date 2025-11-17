package cn.allbs.excel.convert;

import cn.allbs.excel.annotation.NestedProperty;
import cn.allbs.excel.util.NestedFieldResolver;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;

/**
 * 嵌套对象转换器
 * <p>
 * 用于将嵌套对象的字段值提取并导出到 Excel
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty(value = "部门名称", converter = NestedObjectConverter.class)
 * &#64;NestedProperty("dept.name")
 * private Department dept;
 *
 * &#64;ExcelProperty(value = "领导姓名", converter = NestedObjectConverter.class)
 * &#64;NestedProperty(value = "dept.leader.name", nullValue = "暂无")
 * private Department dept;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class NestedObjectConverter implements Converter<Object> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return Object.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 导出时：将嵌套对象转换为 Excel 单元格数据
     * <p>
     * 根据 NestedProperty 注解中的路径表达式提取嵌套字段的值
     * </p>
     */
    @Override
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty,
                                                GlobalConfiguration globalConfiguration) {
        if (value == null) {
            return new WriteCellData<>("");
        }

        try {
            // 获取字段
            Field field = contentProperty.getField();
            if (field == null) {
                log.warn("Field is null, returning original value");
                return new WriteCellData<>(value.toString());
            }

            // 获取 NestedProperty 注解
            NestedProperty nestedProperty = field.getAnnotation(NestedProperty.class);
            if (nestedProperty == null) {
                // 如果没有 NestedProperty 注解，直接返回对象的 toString()
                log.debug("No NestedProperty annotation found on field: {}, using toString()", field.getName());
                return new WriteCellData<>(value.toString());
            }

            // 根据注解配置提取嵌套值
            Object nestedValue = NestedFieldResolver.resolveNestedValue(value, nestedProperty);
            String cellValue = nestedValue != null ? nestedValue.toString() : nestedProperty.nullValue();

            return new WriteCellData<>(cellValue);

        } catch (Exception e) {
            log.error("Failed to convert nested object to excel data", e);
            return new WriteCellData<>(value.toString());
        }
    }

    /**
     * 导入时：暂不支持嵌套对象的导入
     * <p>
     * 嵌套对象的导入需要根据具体业务逻辑来处理，建议使用自定义转换器
     * </p>
     */
    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                     GlobalConfiguration globalConfiguration) {
        log.warn("NestedObjectConverter does not support import. Please use custom converter for nested object import.");
        return cellData.getStringValue();
    }
}
