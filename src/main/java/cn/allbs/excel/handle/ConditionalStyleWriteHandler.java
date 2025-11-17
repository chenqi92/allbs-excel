package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.CellStyleDef;
import cn.allbs.excel.annotation.Condition;
import cn.allbs.excel.annotation.ConditionalStyle;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
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
public class ConditionalStyleWriteHandler implements CellWriteHandler {

	/**
	 * 数据类型
	 */
	private final Class<?> dataClass;

	/**
	 * 列索引 -> 条件样式信息
	 */
	private final Map<Integer, ConditionalStyleInfo> columnStyleMap = new HashMap<>();

	/**
	 * 样式缓存（避免重复创建）
	 */
	private final Map<String, CellStyle> styleCache = new HashMap<>();

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

			if (excelProperty != null && conditionalStyle != null && conditionalStyle.enabled()) {
				// 获取列索引
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				// 排序条件（按优先级）
				Condition[] conditions = conditionalStyle.conditions();
				Arrays.sort(conditions, Comparator.comparingInt(Condition::priority));

				ConditionalStyleInfo styleInfo = new ConditionalStyleInfo(field.getName(), conditions);
				columnStyleMap.put(index, styleInfo);

				columnIndex++;
			}
			else if (excelProperty != null) {
				columnIndex++;
			}
		}

		log.info("Initialized conditional styles for {} columns", columnStyleMap.size());
	}

	@Override
	public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
								  List<WriteCellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
		// 只处理数据行
		if (isHead != null && isHead) {
			return;
		}

		int columnIndex = cell.getColumnIndex();
		ConditionalStyleInfo styleInfo = columnStyleMap.get(columnIndex);

		if (styleInfo == null) {
			return;
		}

		// 获取单元格值
		String cellValue = getCellValueAsString(cell);
		if (cellValue == null) {
			return;
		}

		// 匹配条件并应用样式
		for (Condition condition : styleInfo.conditions) {
			if (matchCondition(cellValue, condition.value())) {
				CellStyle style = getOrCreateStyle(writeSheetHolder.getSheet().getWorkbook(), condition.style());
				cell.setCellStyle(style);
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

		// 设置前景色
		if (!styleDef.foregroundColor().isEmpty()) {
			byte[] rgb = hexToRgb(styleDef.foregroundColor());
			if (rgb != null) {
				cellStyle.setFillForegroundColor(new XSSFColor(rgb, null));
			}
		}

		// 设置背景色
		if (!styleDef.backgroundColor().isEmpty()) {
			byte[] rgb = hexToRgb(styleDef.backgroundColor());
			if (rgb != null) {
				cellStyle.setFillForegroundColor(new XSSFColor(rgb, null));
			}
		}

		// 设置填充模式
		if (styleDef.fillPattern() > 0) {
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
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
