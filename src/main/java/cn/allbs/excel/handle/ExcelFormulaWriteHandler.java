package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelFormula;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel Formula Write Handler
 * <p>
 * Applies Excel formulas to cells based on @ExcelFormula annotations.
 * Supports variable replacement for dynamic formula generation.
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelFormulaWriteHandler implements com.alibaba.excel.write.handler.RowWriteHandler {

	/**
	 * Data class
	 */
	private Class<?> dataClass;

	/**
	 * Column index -> Formula info mapping
	 */
	private final Map<Integer, FormulaInfo> formulaColumnMap = new HashMap<>();

	/**
	 * Default constructor (for reflection instantiation)
	 */
	public ExcelFormulaWriteHandler() {
		this.dataClass = null;
	}

	public ExcelFormulaWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initFormulaColumns();
	}

	/**
	 * Initialize formula column mappings
	 */
	private void initFormulaColumns() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			ExcelFormula excelFormula = field.getAnnotation(ExcelFormula.class);

			if (excelProperty != null) {
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				if (excelFormula != null && excelFormula.enabled()) {
					FormulaInfo info = new FormulaInfo();
					info.field = field;
					info.formula = excelFormula;
					info.columnIndex = index;
					formulaColumnMap.put(index, info);
					log.debug("Registered formula column {}: {} = {}", index, field.getName(), excelFormula.value());
				}

				columnIndex++;
			}
		}

		log.info("Initialized {} formula columns", formulaColumnMap.size());
	}

	// RowWriteHandler methods
	@Override
	public void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
	                             Integer rowIndex, Integer relativeRowIndex, Boolean isHead) {
		// Initialize from data class if not already initialized
		if (dataClass == null && writeSheetHolder.getClazz() != null) {
			this.dataClass = writeSheetHolder.getClazz();
			initFormulaColumns();
			log.info("Initialized ExcelFormulaWriteHandler with data class: {}", dataClass.getSimpleName());
		}
	}

	@Override
	public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
	                            Row row, Integer relativeRowIndex, Boolean isHead) {
		// No action needed
	}

	@Override
	public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
	                             Row row, Integer relativeRowIndex, Boolean isHead) {
		// Skip header row
		if (isHead != null && isHead) {
			return;
		}

		// Skip if no formula columns
		if (formulaColumnMap.isEmpty()) {
			return;
		}

		// Apply formulas to this row
		int rowIndex = row.getRowNum();
		int excelRowNum = rowIndex + 1; // Excel row number (1-based)

		for (Map.Entry<Integer, FormulaInfo> entry : formulaColumnMap.entrySet()) {
			int columnIndex = entry.getKey();
			FormulaInfo formulaInfo = entry.getValue();

			applyFormulaToCell(row, columnIndex, formulaInfo, excelRowNum);
		}
	}

	/**
	 * Apply formula to a single cell
	 */
	private void applyFormulaToCell(Row row, int columnIndex, FormulaInfo formulaInfo, int excelRowNum) {
		ExcelFormula annotation = formulaInfo.formula;
		String formulaTemplate = annotation.value();

		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}

		String columnLetter = CellReference.convertNumToColString(columnIndex);

		// Generate formula for this cell
		String formula = generateFormula(formulaTemplate, excelRowNum, excelRowNum, columnLetter);

		try {
			cell.setCellFormula(formula);
			log.trace("Applied formula to cell [{}, {}]: {}", row.getRowNum(), columnIndex, formula);
		}
		catch (Exception e) {
			log.error("Failed to apply formula to cell [{}, {}]: {}", row.getRowNum(), columnIndex, formula, e);
		}
	}

	/**
	 * Generate formula with variable replacement
	 *
	 * @param template     Formula template
	 * @param excelRow     Excel row number (1-based)
	 * @param lastDataRow  Last data row number
	 * @param columnLetter Current column letter
	 * @return Generated formula
	 */
	private String generateFormula(String template, int excelRow, int lastDataRow, String columnLetter) {
		// Remove leading '=' if present (POI will add it automatically)
		String formula = template.startsWith("=") ? template.substring(1) : template;

		// Replace {row} with current row number
		formula = formula.replace("{row}", String.valueOf(excelRow));

		// Replace {lastRow} with last data row number
		formula = formula.replace("{lastRow}", String.valueOf(lastDataRow));

		// Replace {col} with current column letter
		formula = formula.replace("{col}", columnLetter);

		// Replace {A}, {B}, {C}, etc. with column letters
		for (int i = 0; i < 26; i++) {
			String col = String.valueOf((char) ('A' + i));
			formula = formula.replace("{" + col + "}", col);
		}

		return formula;
	}

	/**
	 * Formula information
	 */
	private static class FormulaInfo {

		Field field;

		ExcelFormula formula;

		int columnIndex;

	}

}
