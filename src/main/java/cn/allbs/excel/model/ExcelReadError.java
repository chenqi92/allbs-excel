package cn.allbs.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Excel Read Error Model
 * <p>
 * Represents an error that occurred during Excel import.
 * Used for collecting errors instead of failing on first error.
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelReadError {

	/**
	 * Sheet index (0-based)
	 */
	private Integer sheetIndex;

	/**
	 * Sheet name
	 */
	private String sheetName;

	/**
	 * Row index (0-based, excluding header)
	 */
	private Integer rowIndex;

	/**
	 * Excel row number (1-based, including header)
	 */
	private Integer excelRowNumber;

	/**
	 * Column index (0-based)
	 */
	private Integer columnIndex;

	/**
	 * Column name (field name or Excel column letter)
	 */
	private String columnName;

	/**
	 * Cell value that caused the error
	 */
	private Object cellValue;

	/**
	 * Error type
	 */
	private ErrorType errorType;

	/**
	 * Error message
	 */
	private String errorMessage;

	/**
	 * Exception that caused the error (optional)
	 */
	private Throwable exception;

	/**
	 * Field name in the data class
	 */
	private String fieldName;

	/**
	 * Expected data type
	 */
	private Class<?> expectedType;

	/**
	 * Error Type Enum
	 */
	public enum ErrorType {
		/**
		 * Data type conversion error
		 */
		TYPE_CONVERSION,

		/**
		 * Validation error (e.g., regex, length, range)
		 */
		VALIDATION,

		/**
		 * Required field is empty
		 */
		REQUIRED_FIELD_EMPTY,

		/**
		 * Invalid format
		 */
		INVALID_FORMAT,

		/**
		 * Duplicate value
		 */
		DUPLICATE_VALUE,

		/**
		 * Unknown error
		 */
		UNKNOWN
	}

	/**
	 * Get a formatted error message for display
	 *
	 * @return Formatted error message
	 */
	public String getFormattedMessage() {
		StringBuilder sb = new StringBuilder();

		if (sheetName != null) {
			sb.append("Sheet [").append(sheetName).append("] ");
		}

		if (excelRowNumber != null) {
			sb.append("Row ").append(excelRowNumber).append(" ");
		}

		if (columnName != null) {
			sb.append("Column [").append(columnName).append("] ");
		}

		sb.append(": ").append(errorMessage);

		if (cellValue != null) {
			sb.append(" (Value: ").append(cellValue).append(")");
		}

		return sb.toString();
	}

}
