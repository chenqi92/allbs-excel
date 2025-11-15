package cn.allbs.excel.convert;

import cn.allbs.excel.annotation.Desensitize;
import cn.allbs.excel.enums.DesensitizeType;
import cn.allbs.excel.util.DesensitizeUtil;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.lang.reflect.Field;

/**
 * 数据脱敏转换器
 * <p>
 * 用于 Excel 导出时对敏感数据进行脱敏处理
 * </p>
 * <p>
 * <strong>注意：</strong>脱敏仅在导出时生效，导入时不进行脱敏处理
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty(value = "手机号", converter = DesensitizeConverter.class)
 * &#64;Desensitize(type = DesensitizeType.MOBILE_PHONE)
 * private String phone;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
public class DesensitizeConverter implements Converter<Object> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return Object.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 导入时：不进行脱敏处理，直接返回原值
     */
    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                    GlobalConfiguration globalConfiguration) {
        return cellData.getStringValue();
    }

    /**
     * 导出时：根据注解配置进行脱敏处理
     */
    @Override
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty,
                                                GlobalConfiguration globalConfiguration) {
        if (value == null) {
            return new WriteCellData<>("");
        }

        Desensitize desensitize = getDesensitize(contentProperty);
        if (desensitize == null || !desensitize.enabled()) {
            return new WriteCellData<>(value.toString());
        }

        String valueStr = value.toString();
        DesensitizeType type = desensitize.type();
        int prefixKeep = desensitize.prefixKeep();
        int suffixKeep = desensitize.suffixKeep();
        String maskChar = desensitize.maskChar();

        // 如果是自定义类型，使用注解中的参数
        if (type == DesensitizeType.CUSTOM) {
            String desensitized = DesensitizeUtil.desensitize(valueStr, type, prefixKeep, suffixKeep, maskChar);
            return new WriteCellData<>(desensitized);
        }

        // 其他类型使用默认参数
        String desensitized = DesensitizeUtil.desensitize(valueStr, type,
                type.getPrefixKeep(), type.getSuffixKeep(), maskChar);
        return new WriteCellData<>(desensitized);
    }

    /**
     * 从字段属性中获取 Desensitize 注解
     */
    private Desensitize getDesensitize(ExcelContentProperty contentProperty) {
        if (contentProperty == null || contentProperty.getField() == null) {
            return null;
        }

        Field field = contentProperty.getField();
        return field.getAnnotation(Desensitize.class);
    }
}

