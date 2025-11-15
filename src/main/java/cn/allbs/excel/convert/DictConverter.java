package cn.allbs.excel.convert;

import cn.allbs.excel.annotation.ExcelDict;
import cn.allbs.excel.service.DictService;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.stream.Collectors;

/**
 * 字典转换器
 * <p>
 * 用于 Excel 导入导出时的字典值转换
 * </p>
 * <ul>
 *     <li>导出时：将字典值转换为字典标签（如：1 -> 男）</li>
 *     <li>导入时：将字典标签转换为字典值（如：男 -> 1）</li>
 * </ul>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty(value = "性别", converter = DictConverter.class)
 * &#64;ExcelDict(dictType = "sys_user_sex")
 * private String sex;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Slf4j
public class DictConverter implements Converter<Object> {

    /**
     * 字典服务，从 Spring 容器中获取
     */
    private static DictService dictService;

    /**
     * 设置字典服务
     * <p>
     * 由 Spring 容器自动注入
     * </p>
     *
     * @param dictService 字典服务
     */
    public static void setDictService(DictService dictService) {
        DictConverter.dictService = dictService;
    }

    @Override
    public Class<?> supportJavaTypeKey() {
        return Object.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 导入时：将 Excel 中的字典标签转换为字典值
     * <p>
     * 例如：将 "男" 转换为 "1"
     * </p>
     */
    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                    GlobalConfiguration globalConfiguration) {
        if (dictService == null) {
            log.warn("DictService 未配置，字典转换功能不可用");
            return cellData.getStringValue();
        }

        ExcelDict excelDict = getExcelDict(contentProperty);
        if (excelDict == null) {
            return cellData.getStringValue();
        }

        String cellValue = cellData.getStringValue();
        if (cellValue == null || cellValue.trim().isEmpty()) {
            return cellValue;
        }

        String dictType = excelDict.dictType();
        String separator = excelDict.separator();

        // 处理多值情况（如：男,女）
        if (cellValue.contains(separator)) {
            return Arrays.stream(cellValue.split(separator))
                    .map(String::trim)
                    .map(label -> dictService.getValue(dictType, label))
                    .filter(value -> value != null && !value.isEmpty())
                    .collect(Collectors.joining(separator));
        }

        // 单值转换
        String dictValue = dictService.getValue(dictType, cellValue.trim());
        return dictValue != null ? dictValue : cellValue;
    }

    /**
     * 导出时：将字典值转换为 Excel 中的字典标签
     * <p>
     * 例如：将 "1" 转换为 "男"
     * </p>
     */
    @Override
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty,
                                                GlobalConfiguration globalConfiguration) {
        if (value == null) {
            return new WriteCellData<>("");
        }

        if (dictService == null) {
            log.warn("DictService 未配置，字典转换功能不可用");
            return new WriteCellData<>(value.toString());
        }

        ExcelDict excelDict = getExcelDict(contentProperty);
        if (excelDict == null) {
            return new WriteCellData<>(value.toString());
        }

        String dictType = excelDict.dictType();
        String separator = excelDict.separator();
        String valueStr = value.toString();

        // 处理多值情况（如：1,2）
        if (valueStr.contains(separator)) {
            String labels = Arrays.stream(valueStr.split(separator))
                    .map(String::trim)
                    .map(val -> dictService.getLabel(dictType, val))
                    .filter(label -> label != null && !label.isEmpty())
                    .collect(Collectors.joining(separator));
            return new WriteCellData<>(labels);
        }

        // 单值转换
        String label = dictService.getLabel(dictType, valueStr.trim());
        return new WriteCellData<>(label != null ? label : valueStr);
    }

    /**
     * 从字段属性中获取 ExcelDict 注解
     */
    private ExcelDict getExcelDict(ExcelContentProperty contentProperty) {
        if (contentProperty == null || contentProperty.getField() == null) {
            return null;
        }

        Field field = contentProperty.getField();
        return field.getAnnotation(ExcelDict.class);
    }
}

