package cn.allbs.excel.exception;

/**
 * Base Excel Exception
 * <p>
 * Base class for all Excel-related exceptions
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
public class ExcelException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	/**
	 * Error code
	 */
	private String errorCode;

	/**
	 * Parameters for error message formatting
	 */
	private Object[] params;

	public ExcelException(String message) {
		super(message);
	}

	public ExcelException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelException(String errorCode, String message) {
		super(message);
		this.errorCode = errorCode;
	}

	public ExcelException(String errorCode, String message, Throwable cause) {
		super(message, cause);
		this.errorCode = errorCode;
	}

	public ExcelException(String errorCode, String message, Object... params) {
		super(String.format(message, params));
		this.errorCode = errorCode;
		this.params = params;
	}

	public ExcelException(String errorCode, String message, Throwable cause, Object... params) {
		super(String.format(message, params), cause);
		this.errorCode = errorCode;
		this.params = params;
	}

	public String getErrorCode() {
		return errorCode;
	}

	public Object[] getParams() {
		return params;
	}

}
