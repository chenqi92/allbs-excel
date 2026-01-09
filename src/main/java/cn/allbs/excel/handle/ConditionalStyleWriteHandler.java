package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.CellStyleDef;
import cn.allbs.excel.annotation.Condition;
import cn.allbs.excel.annotation.ConditionalStyle;
import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.metadata.Head;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.write.handler.CellWriteHandler;
import cn.idev.excel.write.handler.WorkbookWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteSheetHolder;
import cn.idev.excel.write.metadata.holder.WriteTableHolder;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Pattern;

/**
 * 条件样式写处理器
 * <p>
 * 根据单元格值自动应用不同的样式
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class ConditionalStyleWriteHandler implements CellWriteHandler, WorkbookWriteHandler {

	/**
	 * 数据类型
	 */
	private Class<?> dataClass;

	/**
	 * 列索引 -> 条件样式信息
	 */
	private final Map<Integer, ConditionalStyleInfo> columnStyleMap = new HashMap<>();

	/**
	 * 样式缓存（避免重复创建）
	 */
	private final Map<String, CellStyle> styleCache = new HashMap<>();

	/**
	 * Sheet对象（在afterSheetCreate中保存）
	 */
	private Sheet sheet;

	/**
	 * 待应用样式的单元格信息（Sheet名 -> 单元格地址 -> 样式）
	 */
	private final Map<String, Map<CellAddress, CellStyle>> pendingStyles = new HashMap<>();

	/**
	 * 默认构造函数（用于通过反射实例化）
	 */
	public ConditionalStyleWriteHandler() {
		this.dataClass = null;
	}

	public ConditionalStyleWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initConditionalStyles();
	}

	/**
	 * 初始化条件样式
	 */
	private void initConditionalStyles() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			ConditionalStyle conditionalStyle = field.getAnnotation(ConditionalStyle.class);

			// 只处理有 @ExcelProperty 注解的字段
			if (excelProperty != null) {
				// 获取列索引（如果手动指定了 index，使用指定的；否则使用自动递增的）
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				// 如果该字段有条件样式注解且启用
				if (conditionalStyle != null && conditionalStyle.enabled()) {
					// 排序条件（按优先级）
					Condition[] conditions = conditionalStyle.conditions();
					Arrays.sort(conditions, Comparator.comparingInt(Condition::priority));

					ConditionalStyleInfo styleInfo = new ConditionalStyleInfo(field.getName(), conditions);
					columnStyleMap.put(index, styleInfo);

					log.debug("Registered conditional style for column {}: {}", index, field.getName());
				}

				// 每个有 @ExcelProperty 的字段都会占用一列
				columnIndex++;
			}
		}

		log.info("Initialized conditional styles for {} columns", columnStyleMap.size());
	}

	@Override
	public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
								  List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
		// 只处理数据行
		if (isHead != null && isHead) {
			return;
		}

		// 保存sheet对象
		if (this.sheet == null) {
			this.sheet = writeSheetHolder.getSheet();
		}

		int columnIndex = cell.getColumnIndex();
		ConditionalStyleInfo styleInfo = columnStyleMap.get(columnIndex);

		if (styleInfo == null) {
			log.trace("No conditional style for column {}", columnIndex);
			return;
		}

		// 获取单元格值
		String cellValue = getCellValueAsString(cell);
		if (cellValue == null) {
			log.trace("Cell value is null for column {}", columnIndex);
			return;
		}

		log.debug("Processing cell at column {} with value: {}", columnIndex, cellValue);

		// 匹配条件并记录待应用的样式
		for (Condition condition : styleInfo.conditions) {
			if (matchCondition(cellValue, condition.value())) {
				// 创建样式
				CellStyle style = getOrCreateStyle(writeSheetHolder.getSheet().getWorkbook(), condition.style());

				// 记录到待处理列表，在afterWorkbookDispose中统一应用
				String sheetName = writeSheetHolder.getSheet().getSheetName();
				pendingStyles.computeIfAbsent(sheetName, k -> new HashMap<>())
						.put(new CellAddress(cell.getRowIndex(), cell.getColumnIndex()), style);

				log.debug("Scheduled style for cell at column {} (value: {}) matching condition: {}",
						columnIndex, cellValue, condition.value());
				break; // 只应用第一个匹配的条件（优先级最高的）
			}
		}
	}

	/**
	 * 获取或创建样式
	 *
	 * @param workbook 工作簿
	 * @param styleDef 样式定义
	 * @return 单元格样式
	 */
	private CellStyle getOrCreateStyle(Workbook workbook, CellStyleDef styleDef) {
		String cacheKey = generateStyleCacheKey(styleDef);

		return styleCache.computeIfAbsent(cacheKey, key -> createCellStyle(workbook, styleDef));
	}

	/**
	 * 生成样式缓存键
	 *
	 * @param styleDef 样式定义
	 * @return 缓存键
	 */
	private String generateStyleCacheKey(CellStyleDef styleDef) {
		return String.format("%s|%s|%s|%s|%b|%d|%d|%d", styleDef.foregroundColor(), styleDef.backgroundColor(),
				styleDef.fontColor(), styleDef.fillPattern(), styleDef.bold(), styleDef.fontSize(),
				styleDef.horizontalAlignment(), styleDef.verticalAlignment());
	}

	/**
	 * 创建单元格样式
	 *
	 * @param workbook 工作簿
	 * @param styleDef 样式定义
	 * @return 单元格样式
	 */
	private CellStyle createCellStyle(Workbook workbook, CellStyleDef styleDef) {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();

		// 设置背景色（注意：Excel 中背景色使用 FillForegroundColor）
		if (!styleDef.backgroundColor().isEmpty()) {
			byte[] rgb = hexToRgb(styleDef.backgroundColor());
			if (rgb != null) {
				cellStyle.setFillForegroundColor(new XSSFColor(rgb, null));
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		}

		// 设置前景色（单元格图案的前景色，通常不使用）
		if (!styleDef.foregroundColor().isEmpty()) {
			byte[] rgb = hexToRgb(styleDef.foregroundColor());
			if (rgb != null) {
				cellStyle.setFillBackgroundColor(new XSSFColor(rgb, null));
				// 只有在已设置背景色的情况下，前景色才有意义
				if (styleDef.backgroundColor().isEmpty()) {
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
			}
		}

		// 设置字体
		if (!styleDef.fontColor().isEmpty() || styleDef.bold() || styleDef.fontSize() > 0) {
			XSSFFont font = (XSSFFont) workbook.createFont();

			if (!styleDef.fontColor().isEmpty()) {
				byte[] rgb = hexToRgb(styleDef.fontColor());
				if (rgb != null) {
					font.setColor(new XSSFColor(rgb, null));
				}
			}

			if (styleDef.bold()) {
				font.setBold(true);
			}

			if (styleDef.fontSize() > 0) {
				font.setFontHeightInPoints(styleDef.fontSize());
			}

			cellStyle.setFont(font);
		}

		// 设置对齐方式
		if (styleDef.horizontalAlignment() >= 0) {
			cellStyle.setAlignment(HorizontalAlignment.forInt(styleDef.horizontalAlignment()));
		}

		if (styleDef.verticalAlignment() >= 0) {
			cellStyle.setVerticalAlignment(VerticalAlignment.forInt(styleDef.verticalAlignment()));
		}

		return cellStyle;
	}

	/**
	 * 将十六进制颜色转换为 RGB 数组
	 *
	 * @param hexColor 十六进制颜色（如 "#FF0000"）
	 * @return RGB 数组
	 */
	private byte[] hexToRgb(String hexColor) {
		try {
			if (hexColor.startsWith("#")) {
				hexColor = hexColor.substring(1);
			}

			if (hexColor.length() != 6) {
				return null;
			}

			int r = Integer.parseInt(hexColor.substring(0, 2), 16);
			int g = Integer.parseInt(hexColor.substring(2, 4), 16);
			int b = Integer.parseInt(hexColor.substring(4, 6), 16);

			return new byte[] { (byte) r, (byte) g, (byte) b };
		}
		catch (Exception e) {
			log.warn("Invalid hex color: {}", hexColor);
			return null;
		}
	}

	/**
	 * 获取单元格值（字符串形式）
	 *
	 * @param cell 单元格
	 * @return 单元格值
	 */
	private String getCellValueAsString(Cell cell) {
		if (cell == null) {
			return null;
		}

		CellType cellType = cell.getCellType();
		switch (cellType) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue().toString();
				}
				return String.valueOf(cell.getNumericCellValue());
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case FORMULA:
				return cell.getCellFormula();
			default:
				return null;
		}
	}

	/**
	 * 匹配条件
	 *
	 * @param cellValue     单元格值
	 * @param conditionExpr 条件表达式
	 * @return 是否匹配
	 */
	private boolean matchCondition(String cellValue, String conditionExpr) {
		if (cellValue == null || conditionExpr == null) {
			return false;
		}

		try {
			// 1. 正则表达式
			if (conditionExpr.startsWith("regex:")) {
				String regex = conditionExpr.substring(6);
				return Pattern.matches(regex, cellValue);
			}

			// 2. SpEL 表达式（暂不实现，预留接口）
			if (conditionExpr.startsWith("spel:")) {
				log.warn("SpEL expressions are not yet supported: {}", conditionExpr);
				return false;
			}

			// 3. 数值比较
			if (conditionExpr.startsWith(">=")) {
				return compareNumbers(cellValue, conditionExpr.substring(2), ">=");
			}
			if (conditionExpr.startsWith("<=")) {
				return compareNumbers(cellValue, conditionExpr.substring(2), "<=");
			}
			if (conditionExpr.startsWith(">")) {
				return compareNumbers(cellValue, conditionExpr.substring(1), ">");
			}
			if (conditionExpr.startsWith("<")) {
				return compareNumbers(cellValue, conditionExpr.substring(1), "<");
			}

			// 4. 区间
			if ((conditionExpr.startsWith("[") || conditionExpr.startsWith("("))
					&& (conditionExpr.endsWith("]") || conditionExpr.endsWith(")"))) {
				return matchRange(cellValue, conditionExpr);
			}

			// 5. 精确匹配
			return cellValue.equals(conditionExpr);
		}
		catch (Exception e) {
			log.error("Error matching condition: {} with value: {}", conditionExpr, cellValue, e);
			return false;
		}
	}

	/**
	 * 数值比较
	 *
	 * @param cellValue 单元格值
	 * @param threshold 阈值
	 * @param operator  运算符
	 * @return 是否满足条件
	 */
	private boolean compareNumbers(String cellValue, String threshold, String operator) {
		try {
			BigDecimal value = new BigDecimal(cellValue);
			BigDecimal thresholdValue = new BigDecimal(threshold);

			int comparison = value.compareTo(thresholdValue);

			switch (operator) {
				case ">":
					return comparison > 0;
				case ">=":
					return comparison >= 0;
				case "<":
					return comparison < 0;
				case "<=":
					return comparison <= 0;
				default:
					return false;
			}
		}
		catch (NumberFormatException e) {
			return false;
		}
	}

	/**
	 * 区间匹配
	 *
	 * @param cellValue 单元格值
	 * @param rangeExpr 区间表达式（如 "[60,90]"）
	 * @return 是否在区间内
	 */
	private boolean matchRange(String cellValue, String rangeExpr) {
		try {
			boolean leftInclusive = rangeExpr.startsWith("[");
			boolean rightInclusive = rangeExpr.endsWith("]");

			String range = rangeExpr.substring(1, rangeExpr.length() - 1);
			String[] parts = range.split(",");

			if (parts.length != 2) {
				return false;
			}

			BigDecimal value = new BigDecimal(cellValue);
			BigDecimal min = new BigDecimal(parts[0].trim());
			BigDecimal max = new BigDecimal(parts[1].trim());

			boolean minCheck = leftInclusive ? value.compareTo(min) >= 0 : value.compareTo(min) > 0;
			boolean maxCheck = rightInclusive ? value.compareTo(max) <= 0 : value.compareTo(max) < 0;

			return minCheck && maxCheck;
		}
		catch (Exception e) {
			return false;
		}
	}

	@Override
	public void beforeWorkbookCreate() {
		// 无需处理
	}

	@Override
	public void afterWorkbookCreate(WriteWorkbookHolder writeWorkbookHolder) {
		// 无需处理
	}

	@Override
	public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
		// 在工作簿完全写入后，统一应用所有待处理的样式
		if (pendingStyles.isEmpty()) {
			log.debug("No pending styles to apply");
			return;
		}

		log.info("Applying {} pending styles", pendingStyles.values().stream().mapToInt(Map::size).sum());

		Workbook workbook = writeWorkbookHolder.getWorkbook();
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			Map<CellAddress, CellStyle> sheetPendingStyles = pendingStyles.get(sheetName);

			if (sheetPendingStyles == null || sheetPendingStyles.isEmpty()) {
				continue;
			}

			log.debug("Applying {} styles to sheet: {}", sheetPendingStyles.size(), sheetName);

			for (Map.Entry<CellAddress, CellStyle> entry : sheetPendingStyles.entrySet()) {
				CellAddress address = entry.getKey();
				CellStyle style = entry.getValue();

				Row row = sheet.getRow(address.getRow());
				if (row != null) {
					Cell cell = row.getCell(address.getColumn());
					if (cell != null) {
						cell.setCellStyle(style);
						log.trace("Applied style to cell {}", address.formatAsString());
					}
				}
			}
		}

		log.info("Finished applying conditional styles");
	}

	/**
	 * 条件样式信息
	 */
	private static class ConditionalStyleInfo {

		private final String fieldName;

		private final Condition[] conditions;

		public ConditionalStyleInfo(String fieldName, Condition[] conditions) {
			this.fieldName = fieldName;
			this.conditions = conditions;
		}

	}

}
