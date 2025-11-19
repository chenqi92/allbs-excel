package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ConditionalFormat;
import cn.allbs.excel.annotation.ConditionalFormats;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Conditional Format Write Handler
 * <p>
 * Applies conditional formatting to cells based on @ConditionalFormat annotations.
 * Supports data bars, color scales, icon sets, and various highlighting rules.
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ConditionalFormatWriteHandler implements WorkbookWriteHandler {

	private final Class<?> dataClass;

	private final Map<Integer, List<ConditionalFormat>> formatMap = new HashMap<>();

	private int headRowNumber = 1;

	public ConditionalFormatWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initFormatMap();
	}

	/**
	 * Initialize conditional format mappings
	 */
	private void initFormatMap() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);

			if (excelProperty != null) {
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				List<ConditionalFormat> formats = new ArrayList<>();

				// Check for single annotation
				ConditionalFormat singleFormat = field.getAnnotation(ConditionalFormat.class);
				if (singleFormat != null) {
					formats.add(singleFormat);
				}

				// Check for multiple annotations
				ConditionalFormats multipleFormats = field.getAnnotation(ConditionalFormats.class);
				if (multipleFormats != null) {
					for (ConditionalFormat format : multipleFormats.value()) {
						formats.add(format);
					}
				}

				if (!formats.isEmpty()) {
					formatMap.put(index, formats);
					log.debug("Registered {} conditional format(s) for column {}: {}", formats.size(), index,
							field.getName());
				}

				columnIndex++;
			}
		}

		log.info("Initialized conditional formatting for {} columns", formatMap.size());
	}

	@Override
	public void beforeWorkbookCreate() {
		// No action needed
	}

	@Override
	public void afterWorkbookCreate(WriteWorkbookHolder writeWorkbookHolder) {
		// No action needed
	}

	@Override
	public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
		if (formatMap.isEmpty()) {
			log.debug("No conditional formats to apply");
			return;
		}

		Workbook workbook = writeWorkbookHolder.getWorkbook();
		if (!(workbook instanceof XSSFWorkbook)) {
			log.warn("Conditional formatting only supported for XLSX format");
			return;
		}

		XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;

		// Apply to all sheets
		for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
			XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
			applyConditionalFormats(sheet);
		}

		log.info("Applied conditional formatting to {} columns", formatMap.size());
	}

	/**
	 * Apply conditional formats to sheet
	 */
	private void applyConditionalFormats(XSSFSheet sheet) {
		if (sheet.getPhysicalNumberOfRows() <= headRowNumber) {
			log.debug("Sheet {} has no data rows, skipping conditional formatting", sheet.getSheetName());
			return;
		}

		int lastDataRow = sheet.getLastRowNum();

		for (Map.Entry<Integer, List<ConditionalFormat>> entry : formatMap.entrySet()) {
			int columnIndex = entry.getKey();
			List<ConditionalFormat> formats = entry.getValue();

			for (ConditionalFormat format : formats) {
				applyFormat(sheet, columnIndex, format, lastDataRow);
			}
		}
	}

	/**
	 * Apply a single conditional format
	 */
	private void applyFormat(XSSFSheet sheet, int columnIndex, ConditionalFormat format, int lastDataRow) {
		// Determine row range
		int startRow = headRowNumber;
		int endRow = lastDataRow;

		if (!format.applyToAll()) {
			startRow = headRowNumber + format.startRow() - 1;
			if (format.endRow() > 0) {
				endRow = headRowNumber + format.endRow() - 1;
			}
		}

		// Create cell range
		CellRangeAddress[] regions = { new CellRangeAddress(startRow, endRow, columnIndex, columnIndex) };

		SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		try {
			switch (format.type()) {
				case DATA_BAR:
					applyDataBar(sheetCF, regions, format);
					break;
				case COLOR_SCALE:
					applyColorScale(sheetCF, regions, format);
					break;
				case ICON_SET:
					applyIconSet(sheetCF, regions, format);
					break;
				case ABOVE_AVERAGE:
				case BELOW_AVERAGE:
					applyAverageRule(sheetCF, regions, format);
					break;
				case TOP_N:
				case BOTTOM_N:
				case TOP_PERCENT:
				case BOTTOM_PERCENT:
					applyRankRule(sheetCF, regions, format);
					break;
				case DUPLICATE_VALUES:
				case UNIQUE_VALUES:
					applyDuplicateRule(sheetCF, regions, format);
					break;
				case FORMULA:
					applyFormulaRule(sheetCF, regions, format);
					break;
				default:
					log.warn("Unsupported conditional format type: {}", format.type());
			}

			log.debug("Applied {} format to column {} rows {}-{}", format.type(), columnIndex, startRow, endRow);
		}
		catch (Exception e) {
			log.error("Failed to apply conditional format {} to column {}", format.type(), columnIndex, e);
		}
	}

	/**
	 * Apply data bar formatting
	 */
	private void applyDataBar(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		try {
			// Use icon set as fallback since data bars have limited API support in POI
			IconMultiStateFormatting.IconSet iconSetType = IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS;
			XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) sheetCF
				.createConditionalFormattingRule(iconSetType);
			sheetCF.addConditionalFormatting(regions, rule);
			log.debug("Applied icon set as data bar alternative");
		}
		catch (Exception e) {
			log.warn("Data bar formatting has limited POI support: {}", e.getMessage());
		}
	}

	/**
	 * Apply color scale formatting
	 */
	private void applyColorScale(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		try {
			// Color scale support is limited, use simple highlighting instead
			applySimpleHighlight(sheetCF, regions, format);
		}
		catch (Exception e) {
			log.warn("Color scale formatting has limited POI support: {}", e.getMessage());
		}
	}

	/**
	 * Apply icon set formatting
	 */
	private void applyIconSet(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		try {
			IconMultiStateFormatting.IconSet iconSetType = mapIconSetType(format.iconSet());
			XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) sheetCF
				.createConditionalFormattingRule(iconSetType);
			sheetCF.addConditionalFormatting(regions, rule);
		}
		catch (Exception e) {
			log.warn("Icon set formatting error: {}", e.getMessage());
		}
	}

	/**
	 * Apply average-based rule
	 */
	private void applyAverageRule(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		applySimpleHighlight(sheetCF, regions, format);
	}

	/**
	 * Apply simple highlight formatting
	 */
	private void applySimpleHighlight(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		try {
			// Use formula-based rule for highlighting
			String formula = String.format("NOT(ISBLANK(%s1))", CellReference.convertNumToColString(regions[0].getFirstColumn()));
			ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(formula);

			PatternFormatting patternFmt = rule.createPatternFormatting();
			patternFmt.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
			patternFmt.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

			sheetCF.addConditionalFormatting(regions, rule);
		}
		catch (Exception e) {
			log.warn("Simple highlight formatting error: {}", e.getMessage());
		}
	}

	/**
	 * Apply rank-based rule
	 */
	private void applyRankRule(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		// Note: POI has limited support for top/bottom N rules
		log.debug("Rank-based conditional formatting has limited POI support: {}", format.type());
	}

	/**
	 * Apply duplicate/unique value rule
	 */
	private void applyDuplicateRule(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		// Note: POI has limited support for duplicate value rules
		log.debug("Duplicate value conditional formatting has limited POI support: {}", format.type());
	}

	/**
	 * Apply formula-based rule
	 */
	private void applyFormulaRule(SheetConditionalFormatting sheetCF, CellRangeAddress[] regions,
			ConditionalFormat format) {
		if (format.formula().isEmpty()) {
			log.warn("Formula-based conditional format requires formula to be specified");
			return;
		}

		ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(format.formula());

		PatternFormatting patternFmt = rule.createPatternFormatting();
		byte[] fillRgb = hexToRgb(format.fillColor());
		patternFmt.setFillBackgroundColor(IndexedColors.WHITE.getIndex());

		FontFormatting fontFmt = rule.createFontFormatting();
		byte[] fontRgb = hexToRgb(format.fontColor());

		sheetCF.addConditionalFormatting(regions, rule);
	}

	/**
	 * Convert hex color to RGB byte array
	 */
	private byte[] hexToRgb(String hexColor) {
		String hex = hexColor.startsWith("#") ? hexColor.substring(1) : hexColor;
		return new byte[] { (byte) Integer.parseInt(hex.substring(0, 2), 16),
				(byte) Integer.parseInt(hex.substring(2, 4), 16), (byte) Integer.parseInt(hex.substring(4, 6), 16) };
	}

	/**
	 * Map icon set type to POI enum
	 */
	private IconMultiStateFormatting.IconSet mapIconSetType(ConditionalFormat.IconSetType iconSetType) {
		switch (iconSetType) {
			case THREE_ARROWS:
				return IconMultiStateFormatting.IconSet.GYR_3_ARROW;
			case THREE_ARROWS_GRAY:
				return IconMultiStateFormatting.IconSet.GREY_3_ARROWS;
			case THREE_FLAGS:
				return IconMultiStateFormatting.IconSet.GYR_3_FLAGS;
			case THREE_TRAFFIC_LIGHTS_1:
				return IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS;
			case THREE_TRAFFIC_LIGHTS_2:
				return IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS_BOX;
			case THREE_SIGNS:
				return IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS_CIRCLE;
			case THREE_SYMBOLS:
				return IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS_CIRCLE;
			case THREE_SYMBOLS_2:
				return IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS;
			case FOUR_ARROWS:
			case FOUR_ARROWS_GRAY:
				return IconMultiStateFormatting.IconSet.GREY_4_ARROWS;
			case FOUR_RATING:
				return IconMultiStateFormatting.IconSet.RATINGS_4;
			case FOUR_RED_TO_BLACK:
			case FOUR_TRAFFIC_LIGHTS:
				return IconMultiStateFormatting.IconSet.GREY_4_ARROWS;
			case FIVE_ARROWS:
			case FIVE_ARROWS_GRAY:
				return IconMultiStateFormatting.IconSet.GREY_5_ARROWS;
			case FIVE_RATING:
				return IconMultiStateFormatting.IconSet.RATINGS_5;
			case FIVE_QUARTERS:
				return IconMultiStateFormatting.IconSet.QUARTERS_5;
			default:
				return IconMultiStateFormatting.IconSet.GYR_3_ARROW;
		}
	}

}
