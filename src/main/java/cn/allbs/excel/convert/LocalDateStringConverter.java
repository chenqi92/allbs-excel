package cn.allbs.excel.convert;

import cn.allbs.excel.constant.DateConstant;
import cn.allbs.excel.constant.StringPool;
import cn.allbs.excel.util.StringUtil;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;

import java.text.ParseException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

/**
 * 功能: java8 时间转换
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public enum LocalDateStringConverter implements Converter<LocalDate> {
    /**
     * 实例
     */
    INSTANCE;

    @Override
    public Class supportJavaTypeKey() {
        return LocalDate.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public LocalDate convertToJavaData(ReadCellData cellData, ExcelContentProperty contentProperty,
                                       GlobalConfiguration globalConfiguration) throws ParseException {
        if (contentProperty == null || contentProperty.getDateTimeFormatProperty() == null) {
            // 判断时间类型长度
            if (cellData.getStringValue().length() < DateConstant.NORM_DATE_PATTERN.length() && cellData.getStringValue().contains(StringPool.DASH)) {
                String[] time = cellData.getStringValue().split(StringPool.DASH);
                String sb = time[0] +
                        StringPool.DASH +
                        StringUtil.padPre(time[1], 2, "0") +
                        StringPool.DASH +
                        StringUtil.padPre(time[2], 2, "0");
                cellData.setStringValue(sb);
            }
            return LocalDate.parse(cellData.getStringValue());
        } else {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(contentProperty.getDateTimeFormatProperty().getFormat());
            return LocalDate.parse(cellData.getStringValue(), formatter);
        }
    }

    @Override
    public WriteCellData<String> convertToExcelData(LocalDate value, ExcelContentProperty contentProperty,
                                                    GlobalConfiguration globalConfiguration) {
        DateTimeFormatter formatter;
        if (contentProperty == null || contentProperty.getDateTimeFormatProperty() == null) {
            formatter = DateTimeFormatter.ISO_LOCAL_DATE;
        } else {
            formatter = DateTimeFormatter.ofPattern(contentProperty.getDateTimeFormatProperty().getFormat());
        }
        return new WriteCellData<>(value.format(formatter));
    }
}
