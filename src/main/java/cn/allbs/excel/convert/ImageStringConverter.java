package cn.allbs.excel.convert;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
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
        WriteCellData<String> cellData = new WriteCellData<>();
        cellData.setType(CellDataTypeEnum.STRING);
        cellData.setStringValue("");
        return cellData;
    }
}