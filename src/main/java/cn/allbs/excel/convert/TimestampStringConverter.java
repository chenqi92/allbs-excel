package cn.allbs.excel.convert;

import cn.allbs.common.constant.DateConstant;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.sql.Timestamp;
import java.text.ParseException;
import java.time.format.DateTimeFormatter;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public enum TimestampStringConverter implements Converter<Timestamp> {

    /**
     * 实例
     */
    INSTANCE;

    @Override
    public Class supportJavaTypeKey() {
        return Timestamp.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public Timestamp convertToJavaData(ReadCellData cellData, ExcelContentProperty contentProperty,
                                       GlobalConfiguration globalConfiguration) throws ParseException {
        return Timestamp.valueOf(cellData.getStringValue());
    }

    @Override
    public WriteCellData<String> convertToExcelData(Timestamp value, ExcelContentProperty contentProperty,
                                                    GlobalConfiguration globalConfiguration) {
        String timeFormatter = "";
        if (null != value) {
            timeFormatter = value.toLocalDateTime().format(DateTimeFormatter.ofPattern(DateConstant.NORM_DATETIME_PATTERN));
        }
        return new WriteCellData<>(timeFormatter);
    }
}
