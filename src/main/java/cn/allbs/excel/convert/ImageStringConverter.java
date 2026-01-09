package cn.allbs.excel.convert;

import cn.idev.excel.converters.Converter;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

/**
 * 图片字符串转换器
 * <p>
 * 用于处理String类型的图片字段
 * 导入时：保持原样（假设是Base64格式）
 * 导出时：由ImageWriteHandler处理
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class ImageStringConverter implements Converter<String> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return String.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public String convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                   GlobalConfiguration globalConfiguration) {
        // 读取单元格中的字符串（假设是Base64格式的图片数据）
        String stringValue = cellData.getStringValue();
        if (stringValue != null && !stringValue.isEmpty()) {
            // 如果不是Base64格式，添加前缀
            if (!stringValue.startsWith("data:")) {
                // 假设是纯Base64，添加默认的PNG前缀
                // 实际使用时可以根据需要调整
                return "data:image/png;base64," + stringValue;
            }
            return stringValue;
        }
        return null;
    }

    @Override
    public WriteCellData<?> convertToExcelData(String value, ExcelContentProperty contentProperty,
                                               GlobalConfiguration globalConfiguration) {
        // 写入时返回空字符串，实际图片由ImageWriteHandler处理
        // 但需要保留原始值供 ImageWriteHandler 使用
        WriteCellData<String> cellData = new WriteCellData<>();
        cellData.setType(CellDataTypeEnum.STRING);
        cellData.setStringValue("");
        cellData.setData(value); // 保留原始值供 ImageWriteHandler 处理
        return cellData;
    }
}