package cn.allbs.excel.util;

import cn.allbs.excel.annotation.RelatedSheet;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 多 Sheet 关联处理器
 * <p>
 * 用于处理多个 Sheet 之间的关联关系，支持：
 * <ul>
 *     <li>将关联数据导出到不同的 Sheet</li>
 *     <li>在主 Sheet 创建超链接指向关联 Sheet</li>
 *     <li>创建目录 Sheet</li>
 * </ul>
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class MultiSheetRelationProcessor {

	/**
	 * 关联字段信息
	 */
	private static class RelationInfo {

		Field field;

		RelatedSheet annotation;

		int columnIndex;

	}

	/**
	 * 分析数据类，提取关联字段信息
	 *
	 * @param dataClass 数据类
	 * @return 关联字段信息列表
	 */
	public static List<RelationInfo> analyzeRelations(Class<?> dataClass) {
		List<RelationInfo> relationInfos = new ArrayList<>();
		Field[] fields = dataClass.getDeclaredFields();

		for (Field field : fields) {
			RelatedSheet annotation = field.getAnnotation(RelatedSheet.class);
			if (annotation != null && annotation.enabled()) {
				RelationInfo info = new RelationInfo();
				info.field = field;
				info.annotation = annotation;
				relationInfos.add(info);
			}
		}

		return relationInfos;
	}

	/**
	 * 导出多 Sheet 关联数据
	 *
	 * @param writer     ExcelWriter
	 * @param mainData   主数据列表
	 * @param mainSheet  主 Sheet 名称
	 * @param dataClass  数据类型
	 * @param <T>        数据类型
	 */
	public static <T> void exportWithRelations(ExcelWriter writer, List<T> mainData, String mainSheet,
			Class<T> dataClass) {
		List<RelationInfo> relationInfos = analyzeRelations(dataClass);

		if (relationInfos.isEmpty()) {
			log.warn("No @RelatedSheet annotations found in class: {}", dataClass.getName());
			return;
		}

		// 导出主数据到主 Sheet
		WriteSheet mainWriteSheet = com.alibaba.excel.EasyExcel.writerSheet(0, mainSheet).head(dataClass).build();
		writer.write(mainData, mainWriteSheet);

		// 处理每个关联字段
		int sheetIndex = 1;
		for (RelationInfo relationInfo : relationInfos) {
			try {
				processRelatedSheet(writer, mainData, relationInfo, sheetIndex++);
			}
			catch (Exception e) {
				log.error("Failed to process related sheet: {}", relationInfo.annotation.sheetName(), e);
			}
		}
	}

	/**
	 * 处理关联 Sheet
	 */
	private static <T> void processRelatedSheet(ExcelWriter writer, List<T> mainData, RelationInfo relationInfo,
			int sheetIndex) throws Exception {
		Field field = relationInfo.field;
		RelatedSheet annotation = relationInfo.annotation;
		field.setAccessible(true);

		// 收集所有关联数据
		List<Object> allRelatedData = new ArrayList<>();
		Map<Object, Integer> relationKeyToRowMap = new HashMap<>();

		for (int i = 0; i < mainData.size(); i++) {
			T mainItem = mainData.get(i);
			Object relatedData = field.get(mainItem);

			if (relatedData instanceof Collection) {
				Collection<?> collection = (Collection<?>) relatedData;
				for (Object item : collection) {
					allRelatedData.add(item);

					// 记录关联关系（用于后续创建超链接）
					Object relationKeyValue = getRelationKeyValue(mainItem, annotation.relationKey());
					if (relationKeyValue != null) {
						relationKeyToRowMap.put(relationKeyValue, i + 1); // +1 是因为 Excel 行号从 1 开始，且第 1
																			// 行是表头
					}
				}
			}
		}

		if (allRelatedData.isEmpty()) {
			log.debug("No related data found for sheet: {}", annotation.sheetName());
			return;
		}

		// 导出关联数据到新 Sheet
		Class<?> relatedDataType = annotation.dataType();
		if (relatedDataType == Object.class && !allRelatedData.isEmpty()) {
			relatedDataType = allRelatedData.get(0).getClass();
		}

		WriteSheet relatedWriteSheet = com.alibaba.excel.EasyExcel.writerSheet(sheetIndex, annotation.sheetName())
			.head(relatedDataType)
			.build();
		writer.write(allRelatedData, relatedWriteSheet);

		log.info("Exported {} items to related sheet: {}", allRelatedData.size(), annotation.sheetName());
	}

	/**
	 * 获取关联键的值
	 */
	private static Object getRelationKeyValue(Object obj, String fieldName) {
		try {
			Field field = obj.getClass().getDeclaredField(fieldName);
			field.setAccessible(true);
			return field.get(obj);
		}
		catch (Exception e) {
			log.error("Failed to get relation key value: {}", fieldName, e);
			return null;
		}
	}

	/**
	 * 创建超链接（在主 Sheet 中）
	 *
	 * @param workbook        工作簿
	 * @param mainSheet       主 Sheet
	 * @param relatedSheet    关联 Sheet
	 * @param mainRowIndex    主表行索引
	 * @param mainColIndex    主表列索引
	 * @param relatedRowIndex 关联表行索引
	 * @param linkText        链接文本
	 */
	public static void createHyperlink(Workbook workbook, Sheet mainSheet, Sheet relatedSheet, int mainRowIndex,
			int mainColIndex, int relatedRowIndex, String linkText) {
		Row row = mainSheet.getRow(mainRowIndex);
		if (row == null) {
			row = mainSheet.createRow(mainRowIndex);
		}

		Cell cell = row.getCell(mainColIndex);
		if (cell == null) {
			cell = row.createCell(mainColIndex);
		}

		// 创建超链接
		CreationHelper createHelper = workbook.getCreationHelper();
		Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);

		String address = String.format("'%s'!A%d", relatedSheet.getSheetName(), relatedRowIndex + 1);
		hyperlink.setAddress(address);

		cell.setHyperlink(hyperlink);
		cell.setCellValue(linkText);

		// 设置超链接样式
		CellStyle hyperlinkStyle = workbook.createCellStyle();
		Font hyperlinkFont = workbook.createFont();
		hyperlinkFont.setUnderline(Font.U_SINGLE);
		hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
		hyperlinkStyle.setFont(hyperlinkFont);
		cell.setCellStyle(hyperlinkStyle);
	}

	/**
	 * 创建目录 Sheet
	 *
	 * @param workbook       工作簿
	 * @param sheetNames     所有 Sheet 名称
	 * @param indexSheetName 目录 Sheet 名称
	 */
	public static void createIndexSheet(Workbook workbook, List<String> sheetNames, String indexSheetName) {
		Sheet indexSheet = workbook.createSheet(indexSheetName);

		// 移动目录 Sheet 到第一个位置
		workbook.setSheetOrder(indexSheetName, 0);

		// 创建表头
		Row headerRow = indexSheet.createRow(0);
		Cell headerCell1 = headerRow.createCell(0);
		headerCell1.setCellValue("序号");
		Cell headerCell2 = headerRow.createCell(1);
		headerCell2.setCellValue("Sheet 名称");
		Cell headerCell3 = headerRow.createCell(2);
		headerCell3.setCellValue("说明");

		// 设置表头样式
		CellStyle headerStyle = workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 12);
		headerStyle.setFont(headerFont);
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerCell1.setCellStyle(headerStyle);
		headerCell2.setCellStyle(headerStyle);
		headerCell3.setCellStyle(headerStyle);

		// 创建每个 Sheet 的链接
		CreationHelper createHelper = workbook.getCreationHelper();
		for (int i = 0; i < sheetNames.size(); i++) {
			String sheetName = sheetNames.get(i);
			if (sheetName.equals(indexSheetName)) {
				continue;
			}

			Row row = indexSheet.createRow(i + 1);

			// 序号
			Cell cell1 = row.createCell(0);
			cell1.setCellValue(i + 1);

			// Sheet 名称（带超链接）
			Cell cell2 = row.createCell(1);
			Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
			hyperlink.setAddress(String.format("'%s'!A1", sheetName));
			cell2.setHyperlink(hyperlink);
			cell2.setCellValue(sheetName);

			// 设置超链接样式
			CellStyle hyperlinkStyle = workbook.createCellStyle();
			Font hyperlinkFont = workbook.createFont();
			hyperlinkFont.setUnderline(Font.U_SINGLE);
			hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
			hyperlinkStyle.setFont(hyperlinkFont);
			cell2.setCellStyle(hyperlinkStyle);

			// 说明（可以从注解中获取，这里暂时留空）
			Cell cell3 = row.createCell(2);
			cell3.setCellValue("");
		}

		// 自动调整列宽
		indexSheet.autoSizeColumn(0);
		indexSheet.autoSizeColumn(1);
		indexSheet.autoSizeColumn(2);

		log.info("Created index sheet: {}", indexSheetName);
	}

}
