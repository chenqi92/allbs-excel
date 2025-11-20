package cn.allbs.excel.util;

import cn.allbs.excel.annotation.RelatedSheet;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
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
	 * 超链接信息（待创建）
	 */
	public static class HyperlinkInfo {

		public String mainSheetName;

		public String relatedSheetName;

		public int mainRowIndex;

		public int mainColIndex;

		public int relatedRowIndex;

		public String linkText;

		public HyperlinkInfo(String mainSheetName, String relatedSheetName, int mainRowIndex, int mainColIndex,
				int relatedRowIndex, String linkText) {
			this.mainSheetName = mainSheetName;
			this.relatedSheetName = relatedSheetName;
			this.mainRowIndex = mainRowIndex;
			this.mainColIndex = mainColIndex;
			this.relatedRowIndex = relatedRowIndex;
			this.linkText = linkText;
		}

	}

	/**
	 * 分析数据类，提取关联字段信息（带缓存）
	 *
	 * @param dataClass 数据类
	 * @return 关联字段信息列表
	 */
	public static List<RelationInfo> analyzeRelations(Class<?> dataClass) {
		return ClassMetadataCache.getOrCompute(ClassMetadataCache.CacheType.MULTI_SHEET_RELATION, dataClass,
				MultiSheetRelationProcessor::doAnalyzeRelations);
	}

	/**
	 * 实际执行分析（内部方法）
	 */
	private static List<RelationInfo> doAnalyzeRelations(Class<?> dataClass) {
		List<RelationInfo> relationInfos = new ArrayList<>();
		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			RelatedSheet annotation = field.getAnnotation(RelatedSheet.class);

			// 如果有 @RelatedSheet 注解且启用
			if (annotation != null && annotation.enabled()) {
				RelationInfo info = new RelationInfo();
				info.field = field;
				info.annotation = annotation;

				// 如果同时有 @ExcelProperty 注解，使用其 index
				// 否则使用 -1 表示不在主表中显示
				if (excelProperty != null) {
					info.columnIndex = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;
				} else {
					info.columnIndex = -1; // 不在主表中显示
				}

				relationInfos.add(info);
			}

			// 累计列索引
			if (excelProperty != null) {
				columnIndex++;
			}
		}

		log.debug("Analyzed {} related sheet fields for class: {}", relationInfos.size(), dataClass.getName());
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
	 * @return 待创建的超链接列表
	 */
	public static <T> List<HyperlinkInfo> exportWithRelations(ExcelWriter writer, List<T> mainData, String mainSheet,
			Class<T> dataClass) {
		List<HyperlinkInfo> hyperlinks = new ArrayList<>();
		List<RelationInfo> relationInfos = analyzeRelations(dataClass);

		if (relationInfos.isEmpty()) {
			log.warn("No @RelatedSheet annotations found in class: {}", dataClass.getName());
			return hyperlinks;
		}

		// 收集需要排除的字段名（带有 @RelatedSheet 的字段）
		Set<String> excludeFields = relationInfos.stream()
			.map(info -> info.field.getName())
			.collect(java.util.stream.Collectors.toSet());

		// 导出主数据到主 Sheet，排除关联字段
		WriteSheet mainWriteSheet = com.alibaba.excel.EasyExcel.writerSheet(0, mainSheet)
			.head(dataClass)
			.excludeColumnFieldNames(excludeFields)
			.build();
		writer.write(mainData, mainWriteSheet);

		// 处理每个关联字段
		int sheetIndex = 1;
		for (RelationInfo relationInfo : relationInfos) {
			try {
				List<HyperlinkInfo> sheetHyperlinks = processRelatedSheet(writer, mainData, mainSheet, relationInfo,
						sheetIndex++);
				hyperlinks.addAll(sheetHyperlinks);
			}
			catch (Exception e) {
				log.error("Failed to process related sheet: {}", relationInfo.annotation.sheetName(), e);
			}
		}

		return hyperlinks;
	}

	/**
	 * 处理关联 Sheet
	 *
	 * @return 待创建的超链接列表
	 */
	private static <T> List<HyperlinkInfo> processRelatedSheet(ExcelWriter writer, List<T> mainData,
			String mainSheetName, RelationInfo relationInfo, int sheetIndex) throws Exception {
		List<HyperlinkInfo> hyperlinks = new ArrayList<>();
		Field field = relationInfo.field;
		RelatedSheet annotation = relationInfo.annotation;
		field.setAccessible(true);

		// 收集所有关联数据
		List<Object> allRelatedData = new ArrayList<>();
		Map<Object, Integer> relationKeyToRowMap = new HashMap<>();
		Map<Integer, Integer> mainRowToItemCount = new HashMap<>(); // 主表行号 -> 关联数据数量

		for (int i = 0; i < mainData.size(); i++) {
			T mainItem = mainData.get(i);
			Object relatedData = field.get(mainItem);

			if (relatedData instanceof Collection) {
				Collection<?> collection = (Collection<?>) relatedData;
				int itemCount = collection.size();

				if (itemCount > 0) {
					// 记录主表行号对应的关联数据数量
					mainRowToItemCount.put(i, itemCount);

					for (Object item : collection) {
						allRelatedData.add(item);
					}

					// 记录关联关系（用于创建超链接）
					Object relationKeyValue = getRelationKeyValue(mainItem, annotation.relationKey());
					if (relationKeyValue != null) {
						relationKeyToRowMap.put(relationKeyValue, i + 1); // +1 是因为 Excel 行号从 1 开始（0 是表头）
					}
				}
			}
		}

		if (allRelatedData.isEmpty()) {
			log.debug("No related data found for sheet: {}", annotation.sheetName());
			return hyperlinks;
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

		// 创建超链接信息（如果启用）
		if (annotation.createHyperlink()) {
			for (Map.Entry<Integer, Integer> entry : mainRowToItemCount.entrySet()) {
				int mainRowIndex = entry.getKey() + 1; // Excel 行号（0 是表头，数据从第 1 行开始）
				int itemCount = entry.getValue();

				// 生成超链接文本
				String linkText = annotation.hyperlinkText();
				if (linkText.isEmpty()) {
					linkText = String.format("查看明细(%d)", itemCount);
				}

				// 创建超链接信息
				HyperlinkInfo hyperlinkInfo = new HyperlinkInfo(mainSheetName, annotation.sheetName(), mainRowIndex,
						relationInfo.columnIndex, 1, // 跳转到关联表的第 1 行（表头下方）
						linkText);
				hyperlinks.add(hyperlinkInfo);
			}

			log.debug("Created {} hyperlinks for sheet: {}", hyperlinks.size(), annotation.sheetName());
		}

		return hyperlinks;
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
	 * 批量应用超链接
	 *
	 * @param workbook  工作簿
	 * @param hyperlinks 超链接信息列表
	 */
	public static void applyHyperlinks(Workbook workbook, List<HyperlinkInfo> hyperlinks) {
		if (hyperlinks == null || hyperlinks.isEmpty()) {
			log.debug("No hyperlinks to apply");
			return;
		}

		// 创建超链接样式（只创建一次，复用）
		CellStyle hyperlinkStyle = workbook.createCellStyle();
		Font hyperlinkFont = workbook.createFont();
		hyperlinkFont.setUnderline(Font.U_SINGLE);
		hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
		hyperlinkStyle.setFont(hyperlinkFont);

		CreationHelper createHelper = workbook.getCreationHelper();

		for (HyperlinkInfo info : hyperlinks) {
			try {
				Sheet mainSheet = workbook.getSheet(info.mainSheetName);
				if (mainSheet == null) {
					log.warn("Main sheet not found: {}", info.mainSheetName);
					continue;
				}

				Row row = mainSheet.getRow(info.mainRowIndex);
				if (row == null) {
					row = mainSheet.createRow(info.mainRowIndex);
				}

				Cell cell = row.getCell(info.mainColIndex);
				if (cell == null) {
					cell = row.createCell(info.mainColIndex);
				}

				// 创建超链接
				Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
				String address = String.format("'%s'!A%d", info.relatedSheetName, info.relatedRowIndex + 1);
				hyperlink.setAddress(address);

				cell.setHyperlink(hyperlink);
				cell.setCellValue(info.linkText);
				cell.setCellStyle(hyperlinkStyle);

				log.trace("Created hyperlink at [{}, {}]: {}", info.mainRowIndex, info.mainColIndex, info.linkText);
			}
			catch (Exception e) {
				log.error("Failed to create hyperlink at [{}, {}]", info.mainRowIndex, info.mainColIndex, e);
			}
		}

		log.info("Applied {} hyperlinks", hyperlinks.size());
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
	 * @deprecated 使用 {@link #applyHyperlinks(Workbook, List)} 代替
	 */
	@Deprecated
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
