package cn.allbs.excel.exception;

/**
 * Excel Import Exception
 * <p>
 * Thrown when errors occur during Excel import
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
public class ExcelImportException extends ExcelException {

	private static final long serialVersionUID = 1L;

	/**
	 * Sheet name where error occurred
	 */
	private String sheetName;

	/**
	 * Row index where error occurred (0-based, excluding header)
	 */
	private Integer rowIndex;

	/**
	 * Excel row number (1-based, including header)
	 */
	private Integer excelRowNumber;

	/**
	 * Column index where error occurred (0-based)
	 */
	private Integer columnIndex;

	/**
	 * Column name (header name)
	 */
	private String columnName;

	/**
	 * Field name in data class
	 */
	private String fieldName;

	/**
	 * Cell value that caused the error
	 */
	private Object cellValue;

	/**
	 * Expected data type
	 */
	private Class<?> expectedType;

	/**
	 * Actual value type
	 */
	private Class<?> actualType;

	public ExcelImportException(String message) {
		super(message);
	}

	public ExcelImportException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelImportException(String errorCode, String message) {
		super(errorCode, message);
	}

	public ExcelImportException(String errorCode, String message, Throwable cause) {
		super(errorCode, message, cause);
	}

	public ExcelImportException(String errorCode, String message, Object... params) {
		super(errorCode, message, params);
	}

	public ExcelImportException(String errorCode, String message, Throwable cause, Object... params) {
		super(errorCode, message, cause, params);
	}

	// Builder pattern for more detailed exceptions
	public static Builder builder() {
		return new Builder();
	}

	public static class Builder {

		private String errorCode;

		private String message;

		private Throwable cause;

		private String sheetName;

		private Integer rowIndex;

		private Integer excelRowNumber;

		private Integer columnIndex;

		private String columnName;

		private String fieldName;

		private Object cellValue;

		private Class<?> expectedType;

		private Class<?> actualType;

		private Object[] params;

		public Builder errorCode(String errorCode) {
			this.errorCode = errorCode;
			return this;
		}

		public Builder message(String message) {
			this.message = message;
			return this;
		}

		public Builder cause(Throwable cause) {
			this.cause = cause;
			return this;
		}

		public Builder sheetName(String sheetName) {
			this.sheetName = sheetName;
			return this;
		}

		public Builder rowIndex(Integer rowIndex) {
			this.rowIndex = rowIndex;
			return this;
		}

		public Builder excelRowNumber(Integer excelRowNumber) {
			this.excelRowNumber = excelRowNumber;
			return this;
		}

		public Builder columnIndex(Integer columnIndex) {
			this.columnIndex = columnIndex;
			return this;
		}

		public Builder columnName(String columnName) {
			this.columnName = columnName;
			return this;
		}

		public Builder fieldName(String fieldName) {
			this.fieldName = fieldName;
			return this;
		}

		public Builder cellValue(Object cellValue) {
			this.cellValue = cellValue;
			return this;
		}

		public Builder expectedType(Class<?> expectedType) {
			this.expectedType = expectedType;
			return this;
		}

		public Builder actualType(Class<?> actualType) {
			this.actualType = actualType;
			return this;
		}

		public Builder params(Object... params) {
			this.params = params;
			return this;
		}

		public ExcelImportException build() {
			ExcelImportException exception;
			if (cause != null) {
				exception = params != null ? new ExcelImportException(errorCode, message, cause, params)
						: new ExcelImportException(errorCode, message, cause);
			}
			else {
				exception = params != null ? new ExcelImportException(errorCode, message, params)
						: new ExcelImportException(errorCode, message);
			}

			exception.sheetName = this.sheetName;
			exception.rowIndex = this.rowIndex;
			exception.excelRowNumber = this.excelRowNumber;
			exception.columnIndex = this.columnIndex;
			exception.columnName = this.columnName;
			exception.fieldName = this.fieldName;
			exception.cellValue = this.cellValue;
			exception.expectedType = this.expectedType;
			exception.actualType = this.actualType;

			return exception;
		}

	}

	// Getters
	public String getSheetName() {
		return sheetName;
	}

	public Integer getRowIndex() {
		return rowIndex;
	}

	public Integer getExcelRowNumber() {
		return excelRowNumber;
	}

	public Integer getColumnIndex() {
		return columnIndex;
	}

	public String getColumnName() {
		return columnName;
	}

	public String getFieldName() {
		return fieldName;
	}

	public Object getCellValue() {
		return cellValue;
	}

	public Class<?> getExpectedType() {
		return expectedType;
	}

	public Class<?> getActualType() {
		return actualType;
	}

	/**
	 * Get formatted error message with context
	 */
	public String getDetailedMessage() {
		StringBuilder sb = new StringBuilder(getMessage());

		if (sheetName != null) {
			sb.append(" [Sheet: ").append(sheetName).append("]");
		}
		if (excelRowNumber != null) {
			sb.append(" [Row: ").append(excelRowNumber).append("]");
		}
		if (columnName != null) {
			sb.append(" [Column: ").append(columnName).append("]");
		}
		else if (columnIndex != null) {
			sb.append(" [Column Index: ").append(columnIndex).append("]");
		}
		if (fieldName != null) {
			sb.append(" [Field: ").append(fieldName).append("]");
		}
		if (cellValue != null) {
			sb.append(" [Value: ").append(cellValue).append("]");
		}
		if (expectedType != null && actualType != null) {
			sb.append(" [Expected: ")
				.append(expectedType.getSimpleName())
				.append(", Actual: ")
				.append(actualType.getSimpleName())
				.append("]");
		}

		return sb.toString();
	}

}
