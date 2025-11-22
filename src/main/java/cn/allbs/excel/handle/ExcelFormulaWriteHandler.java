package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelFormula;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
public class ExcelFormulaWriteHandler implements SheetWriteHandler, WorkbookWriteHandler {

	/**
	 * Data class
	 */
	private Class<?> dataClass;

	/**
	 * Column index -> Formula info mapping
	 */
	private final Map<Integer, FormulaInfo> formulaColumnMap = new HashMap<>();

	/**
	 * Number of header rows
	 */
	private int headRowNumber = 1;

	/**
	 * Number of data rows
	 */
	private int dataRowCount = 0;

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

	// SheetWriteHandler methods
	@Override
	public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
		// Initialize from data class if not already initialized
		if (dataClass == null && writeSheetHolder.getClazz() != null) {
			this.dataClass = writeSheetHolder.getClazz();
			initFormulaColumns();
			log.info("Initialized ExcelFormulaWriteHandler with data class: {}", dataClass.getSimpleName());
		}
	}

	@Override
	public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
		// No action needed - formulas will be applied in afterWorkbookDispose
		log.debug("Sheet created, formulas will be applied after all data is written");
	}

	// WorkbookWriteHandler methods
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
		// Apply formulas after all data has been written
		if (formulaColumnMap.isEmpty()) {
			log.debug("No formula columns to process");
			return;
		}

		Workbook workbook = writeWorkbookHolder.getWorkbook();

		// Process all sheets
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			applyFormulas(sheet);
		}

		log.info("Applied formulas to {} columns", formulaColumnMap.size());
	}

	/**
	 * Apply formulas to sheet
	 */
	private void applyFormulas(Sheet sheet) {
		if (sheet.getPhysicalNumberOfRows() <= headRowNumber) {
			log.debug("Sheet {} has no data rows, skipping formula application", sheet.getSheetName());
			return;
		}

		// Calculate data row count
		dataRowCount = sheet.getLastRowNum() - headRowNumber + 1;
		int lastDataRow = sheet.getLastRowNum() + 1; // Excel row number (1-based)

		log.debug("Applying formulas to sheet: {}, data rows: {}, last row: {}", sheet.getSheetName(), dataRowCount,
				lastDataRow);

		// Apply formulas to each column
		for (Map.Entry<Integer, FormulaInfo> entry : formulaColumnMap.entrySet()) {
			int columnIndex = entry.getKey();
			FormulaInfo formulaInfo = entry.getValue();

			applyFormulaToColumn(sheet, columnIndex, formulaInfo, lastDataRow);
		}
	}

	/**
	 * Apply formula to a column
	 */
	private void applyFormulaToColumn(Sheet sheet, int columnIndex, FormulaInfo formulaInfo, int lastDataRow) {
		ExcelFormula annotation = formulaInfo.formula;
		String formulaTemplate = annotation.value();

		// Determine row range
		int startRow = headRowNumber;
		int endRow = sheet.getLastRowNum();

		if (annotation.limitToRange()) {
			startRow = headRowNumber + annotation.startRow() - 1;
			if (annotation.endRow() > 0) {
				endRow = headRowNumber + annotation.endRow() - 1;
			}
		}

		String columnLetter = CellReference.convertNumToColString(columnIndex);

		// Apply formula to each row
		for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}

			Cell cell = row.getCell(columnIndex);
			if (cell == null) {
				cell = row.createCell(columnIndex);
			}

			// Generate formula for this cell
			String formula = generateFormula(formulaTemplate, rowIndex + 1, // Excel row number (1-based)
					lastDataRow, columnLetter);

			try {
				cell.setCellFormula(formula);
				log.trace("Applied formula to cell [{}, {}]: {}", rowIndex, columnIndex, formula);
			}
			catch (Exception e) {
				log.error("Failed to apply formula to cell [{}, {}]: {}", rowIndex, columnIndex, formula, e);
			}
		}

		log.debug("Applied formula to column {}: {} cells", columnIndex, endRow - startRow + 1);
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
