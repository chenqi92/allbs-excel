package cn.allbs.excel.convert;

import cn.idev.excel.converters.Converter;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;

import java.util.List;

/**
 * 图片列表转换器
 * <p>
 * 用于处理 List 类型的图片字段，将列表转换为空字符串写入单元格
 * 实际的图片渲染由 ImageWriteHandler 处理
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
public class ImageListConverter implements Converter<List> {

	@Override
	public Class<?> supportJavaTypeKey() {
		return List.class;
	}

	@Override
	public CellDataTypeEnum supportExcelTypeKey() {
		return CellDataTypeEnum.STRING;
	}

	@Override
	public List convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
			GlobalConfiguration globalConfiguration) {
		// 读取时不需要转换
		return null;
	}

	@Override
	public WriteCellData<?> convertToExcelData(List value, ExcelContentProperty contentProperty,
			GlobalConfiguration globalConfiguration) {
		// 写入时返回空字符串，图片由 ImageWriteHandler 处理
		// 保留原始数据在 WriteCellData 中，以便 ImageWriteHandler 可以访问
		WriteCellData<List> cellData = new WriteCellData<>();
		cellData.setType(CellDataTypeEnum.STRING);
		cellData.setStringValue("");
		cellData.setData(value); // 保存原始 List 数据
		return cellData;
	}

}
