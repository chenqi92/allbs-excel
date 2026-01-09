package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelComment;
import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.write.handler.WorkbookWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel Comment Write Handler
 * <p>
 * Adds comments to Excel cells based on @ExcelComment annotations.
 * Comments appear as small indicators and show text on hover.
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelCommentWriteHandler implements WorkbookWriteHandler, cn.idev.excel.write.handler.SheetWriteHandler {

	private Class<?> dataClass;

	private final Map<Integer, CommentInfo> commentMap = new HashMap<>();

	private int headRowNumber = 0; // Header row index (0-based)

	/**
	 * Default constructor (for reflection instantiation)
	 */
	public ExcelCommentWriteHandler() {
		this.dataClass = null;
	}

	public ExcelCommentWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initCommentMap();
	}

	/**
	 * Initialize comment mappings
	 */
	private void initCommentMap() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			ExcelComment excelComment = field.getAnnotation(ExcelComment.class);

			if (excelProperty != null && excelComment != null) {
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				CommentInfo info = new CommentInfo();
				info.field = field;
				info.annotation = excelComment;
				info.columnIndex = index;

				commentMap.put(index, info);
				log.debug("Registered comment for column {}: {}", index, field.getName());

				columnIndex++;
			}
			else if (excelProperty != null) {
				columnIndex++;
			}
		}

		log.info("Initialized comments for {} columns", commentMap.size());
	}

	// SheetWriteHandler methods
	@Override
	public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, cn.idev.excel.write.metadata.holder.WriteSheetHolder writeSheetHolder) {
		// Initialize from data class if not already initialized
		if (dataClass == null && writeSheetHolder.getClazz() != null) {
			this.dataClass = writeSheetHolder.getClazz();
			initCommentMap();
			log.info("Initialized ExcelCommentWriteHandler with data class: {}", dataClass.getSimpleName());
		}
	}

	@Override
	public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, cn.idev.excel.write.metadata.holder.WriteSheetHolder writeSheetHolder) {
		// No action needed - comments will be applied in afterWorkbookDispose
		log.debug("Sheet created, comments will be applied after all data is written");
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
		if (commentMap.isEmpty()) {
			log.debug("No comments to apply");
			return;
		}

		Workbook workbook = writeWorkbookHolder.getWorkbook();

		// Apply to all sheets
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			applyComments(sheet);
		}

		log.info("Applied comments to {} columns", commentMap.size());
	}

	/**
	 * Apply comments to sheet
	 */
	private void applyComments(Sheet sheet) {
		if (!(sheet instanceof XSSFSheet)) {
			log.warn("Comments only fully supported for XLSX format");
			return;
		}

		XSSFSheet xssfSheet = (XSSFSheet) sheet;
		CreationHelper factory = sheet.getWorkbook().getCreationHelper();
		Drawing<?> drawing = xssfSheet.createDrawingPatriarch();

		for (Map.Entry<Integer, CommentInfo> entry : commentMap.entrySet()) {
			int columnIndex = entry.getKey();
			CommentInfo commentInfo = entry.getValue();

			// Add header comment
			if (!commentInfo.annotation.headerComment().isEmpty()) {
				addComment(xssfSheet, drawing, factory, headRowNumber, columnIndex, commentInfo.annotation.headerComment(),
						commentInfo.annotation);
			}

			// Add data comments
			if (!commentInfo.annotation.dataComment().isEmpty()) {
				addDataComments(xssfSheet, drawing, factory, columnIndex, commentInfo);
			}
		}
	}

	/**
	 * Add data comments to column
	 */
	private void addDataComments(XSSFSheet sheet, Drawing<?> drawing, CreationHelper factory, int columnIndex,
			CommentInfo commentInfo) {
		ExcelComment annotation = commentInfo.annotation;

		// Determine row range
		int startRow = headRowNumber + 1; // First data row
		int endRow = sheet.getLastRowNum();

		if (!annotation.applyToAll()) {
			startRow = headRowNumber + annotation.startRow();
			if (annotation.endRow() > 0) {
				endRow = headRowNumber + annotation.endRow();
			}
		}

		// Add comment to each row
		for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}

			Cell cell = row.getCell(columnIndex);
			if (cell == null) {
				continue;
			}

			// Generate comment text with placeholders
			String commentText = generateCommentText(annotation.dataComment(), cell, rowIndex, columnIndex);

			addComment(sheet, drawing, factory, rowIndex, columnIndex, commentText, annotation);
		}
	}

	/**
	 * Add comment to a cell
	 */
	private void addComment(XSSFSheet sheet, Drawing<?> drawing, CreationHelper factory, int rowIndex, int columnIndex,
			String commentText, ExcelComment annotation) {
		// Create anchor for comment box
		ClientAnchor anchor = factory.createClientAnchor();
		anchor.setCol1(columnIndex + 1); // Start one column to the right
		anchor.setCol2(columnIndex + 1 + annotation.width());
		anchor.setRow1(rowIndex);
		anchor.setRow2(rowIndex + annotation.height());

		// Create comment
		Comment comment = drawing.createCellComment(anchor);
		RichTextString richText = factory.createRichTextString(commentText);
		comment.setString(richText);
		comment.setAuthor(annotation.author());
		comment.setVisible(annotation.visible());

		// Attach comment to cell
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}

		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}

		cell.setCellComment(comment);

		log.trace("Added comment to cell [{}, {}]: {}", rowIndex, columnIndex, commentText);
	}

	/**
	 * Generate comment text with placeholder replacement
	 */
	private String generateCommentText(String template, Cell cell, int rowIndex, int columnIndex) {
		String text = template;

		// Replace {value} with cell value
		String cellValue = getCellValueAsString(cell);
		text = text.replace("{value}", cellValue);

		// Replace {row} with Excel row number (1-based)
		text = text.replace("{row}", String.valueOf(rowIndex + 1));

		// Replace {col} with column letter
		text = text.replace("{col}", CellReference.convertNumToColString(columnIndex));

		return text;
	}

	/**
	 * Get cell value as string
	 */
	private String getCellValueAsString(Cell cell) {
		if (cell == null) {
			return "";
		}

		switch (cell.getCellType()) {
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
			case BLANK:
				return "";
			default:
				return cell.toString();
		}
	}

	/**
	 * Comment information
	 */
	private static class CommentInfo {

		Field field;

		ExcelComment annotation;

		int columnIndex;

	}

}
