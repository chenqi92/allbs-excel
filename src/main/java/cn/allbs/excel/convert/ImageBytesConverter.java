package cn.allbs.excel.convert;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

import java.util.Base64;

/**
 * 图片字节数组转换器
 * <p>
 * 用于处理 byte[] 类型的图片字段
 * 导入时：将Base64字符串转换为byte[]
 * 导出时：将byte[]转换为Base64字符串
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class ImageBytesConverter implements Converter<byte[]> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return byte[].class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public byte[] convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                   GlobalConfiguration globalConfiguration) {
        try {
            String stringValue = cellData.getStringValue();
            if (stringValue == null || stringValue.isEmpty()) {
                return null;
            }

            // 处理Base64格式的图片数据
            if (stringValue.startsWith("data:")) {
                // 移除data:image/png;base64,前缀
                int commaIndex = stringValue.indexOf(",");
                if (commaIndex > 0) {
                    String base64Only = stringValue.substring(commaIndex + 1);
                    return Base64.getDecoder().decode(base64Only);
                }
            } else {
                // 直接的Base64字符串
                return Base64.getDecoder().decode(stringValue);
            }
        } catch (Exception e) {
            log.error("转换图片数据失败: {}", e.getMessage());
        }
        return null;
    }

    @Override
    public WriteCellData<?> convertToExcelData(byte[] value, ExcelContentProperty contentProperty,
                                               GlobalConfiguration globalConfiguration) {
        // 写入时返回空字符串，实际图片由ImageWriteHandler处理
        WriteCellData<String> cellData = new WriteCellData<>();
        cellData.setType(CellDataTypeEnum.STRING);
        cellData.setStringValue("");
        return cellData;
    }
}