package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelSheetStyle;
import cn.idev.excel.write.handler.SheetWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteSheetHolder;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excel Sheet Style Write Handler
 * <p>
 * Applies sheet-level styles based on @ExcelSheetStyle annotation
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelSheetStyleWriteHandler implements SheetWriteHandler {

	private Class<?> dataClass;

	private ExcelSheetStyle styleAnnotation;

	/**
	 * Default constructor (for reflection instantiation)
	 */
	public ExcelSheetStyleWriteHandler() {
		this.dataClass = null;
	}

	public ExcelSheetStyleWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initStyleAnnotation();
	}

	/**
	 * Initialize style annotation
	 */
	private void initStyleAnnotation() {
		if (dataClass != null) {
			styleAnnotation = dataClass.getAnnotation(ExcelSheetStyle.class);
			if (styleAnnotation != null) {
				log.debug("Found @ExcelSheetStyle annotation on class: {}", dataClass.getName());
			}
		}
	}

	@Override
	public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
		// Initialize from data class if not already initialized
		if (dataClass == null && writeSheetHolder.getClazz() != null) {
			this.dataClass = writeSheetHolder.getClazz();
			initStyleAnnotation();
			log.info("Initialized ExcelSheetStyleWriteHandler with data class: {}", dataClass.getSimpleName());
		}
	}

	@Override
	public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
		// Try lazy initialization if not initialized in beforeSheetCreate
		if (dataClass == null && writeSheetHolder.getClazz() != null) {
			this.dataClass = writeSheetHolder.getClazz();
			initStyleAnnotation();
			log.info("Late initialization of ExcelSheetStyleWriteHandler with data class: {}",
					dataClass.getSimpleName());
		}

		if (styleAnnotation == null) {
			log.debug("No @ExcelSheetStyle annotation found, skipping sheet style application");
			return;
		}

		Sheet sheet = writeSheetHolder.getSheet();
		applySheetStyles(sheet);
	}

	/**
	 * Apply sheet styles
	 */
	private void applySheetStyles(Sheet sheet) {
		// Apply freeze pane
		if (styleAnnotation.freezeRow() > 0 || styleAnnotation.freezeColumn() > 0) {
			applyFreezePane(sheet);
		}

		// Apply auto-filter
		if (styleAnnotation.autoFilter()) {
			applyAutoFilter(sheet);
		}

		// Apply default column width
		if (styleAnnotation.defaultColumnWidth() > 0) {
			sheet.setDefaultColumnWidth(styleAnnotation.defaultColumnWidth());
			log.debug("Set default column width: {}", styleAnnotation.defaultColumnWidth());
		}

		// Apply zoom scale
		if (styleAnnotation.zoomScale() >= 10 && styleAnnotation.zoomScale() <= 400) {
			sheet.setZoom(styleAnnotation.zoomScale());
			log.debug("Set zoom scale: {}%", styleAnnotation.zoomScale());
		}

		// Apply grid lines visibility
		sheet.setDisplayGridlines(styleAnnotation.showGridLines());
	}

	/**
	 * Apply freeze pane
	 */
	private void applyFreezePane(Sheet sheet) {
		int freezeRow = styleAnnotation.freezeRow();
		int freezeColumn = styleAnnotation.freezeColumn();

		// Create freeze pane
		// Parameters: colSplit, rowSplit, leftmostColumn, topRow
		sheet.createFreezePane(freezeColumn, freezeRow, freezeColumn, freezeRow);

		log.info("Applied freeze pane: row={}, column={}", freezeRow, freezeColumn);
	}

	/**
	 * Apply auto-filter
	 * Note: Auto-filter will be applied with a default large range.
	 * The actual range will be automatically adjusted by Excel when data is present.
	 */
	private void applyAutoFilter(Sheet sheet) {
		int firstRow = styleAnnotation.autoFilterStartRow();
		// Set auto-filter on a large range
		// Excel will automatically adjust it based on actual data
		CellRangeAddress filterRange = new CellRangeAddress(firstRow, 65535, 0, 255);
		sheet.setAutoFilter(filterRange);
		log.info("Applied auto-filter starting from row: {}", firstRow);
	}

}
