package cn.allbs.excel.exception;

/**
 * Excel Export Exception
 * <p>
 * Thrown when errors occur during Excel export
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
public class ExcelExportException extends ExcelException {

	private static final long serialVersionUID = 1L;

	/**
	 * Sheet name where error occurred
	 */
	private String sheetName;

	/**
	 * Row index where error occurred (0-based)
	 */
	private Integer rowIndex;

	/**
	 * Column index where error occurred (0-based)
	 */
	private Integer columnIndex;

	/**
	 * Field name that caused the error
	 */
	private String fieldName;

	public ExcelExportException(String message) {
		super(message);
	}

	public ExcelExportException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelExportException(String errorCode, String message) {
		super(errorCode, message);
	}

	public ExcelExportException(String errorCode, String message, Throwable cause) {
		super(errorCode, message, cause);
	}

	public ExcelExportException(String errorCode, String message, Object... params) {
		super(errorCode, message, params);
	}

	public ExcelExportException(String errorCode, String message, Throwable cause, Object... params) {
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

		private Integer columnIndex;

		private String fieldName;

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

		public Builder columnIndex(Integer columnIndex) {
			this.columnIndex = columnIndex;
			return this;
		}

		public Builder fieldName(String fieldName) {
			this.fieldName = fieldName;
			return this;
		}

		public Builder params(Object... params) {
			this.params = params;
			return this;
		}

		public ExcelExportException build() {
			ExcelExportException exception;
			if (cause != null) {
				exception = params != null ? new ExcelExportException(errorCode, message, cause, params)
						: new ExcelExportException(errorCode, message, cause);
			}
			else {
				exception = params != null ? new ExcelExportException(errorCode, message, params)
						: new ExcelExportException(errorCode, message);
			}

			exception.sheetName = this.sheetName;
			exception.rowIndex = this.rowIndex;
			exception.columnIndex = this.columnIndex;
			exception.fieldName = this.fieldName;

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

	public Integer getColumnIndex() {
		return columnIndex;
	}

	public String getFieldName() {
		return fieldName;
	}

	/**
	 * Get formatted error message with context
	 */
	public String getDetailedMessage() {
		StringBuilder sb = new StringBuilder(getMessage());

		if (sheetName != null) {
			sb.append(" [Sheet: ").append(sheetName).append("]");
		}
		if (rowIndex != null) {
			sb.append(" [Row: ").append(rowIndex + 1).append("]"); // 1-based for user display
		}
		if (columnIndex != null) {
			sb.append(" [Column: ").append(columnIndex).append("]");
		}
		if (fieldName != null) {
			sb.append(" [Field: ").append(fieldName).append("]");
		}

		return sb.toString();
	}

}
