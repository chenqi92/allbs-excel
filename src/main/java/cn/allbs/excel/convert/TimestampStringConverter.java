package cn.allbs.excel.convert;

import cn.allbs.excel.constant.DateConstant;
import org.apache.fesod.sheet.converters.Converter;
import org.apache.fesod.sheet.enums.CellDataTypeEnum;
import org.apache.fesod.sheet.metadata.GlobalConfiguration;
import org.apache.fesod.sheet.metadata.data.ReadCellData;
import org.apache.fesod.sheet.metadata.data.WriteCellData;
import org.apache.fesod.sheet.metadata.property.ExcelContentProperty;

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
