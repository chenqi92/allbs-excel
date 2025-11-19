package cn.allbs.excel.constant;

/**
 * Excel Constants
 * <p>
 * Centralized constants for Excel operations
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
public final class ExcelConstants {

	private ExcelConstants() {
		throw new UnsupportedOperationException("Utility class cannot be instantiated");
	}

	// ==================== Excel Format Limits ====================

	/**
	 * Maximum rows in .xls format (Excel 97-2003)
	 */
	public static final int MAX_ROWS_XLS = 65535;

	/**
	 * Maximum rows in .xlsx format (Excel 2007+)
	 */
	public static final int MAX_ROWS_XLSX = 1048576;

	/**
	 * Maximum columns in .xls format
	 */
	public static final int MAX_COLUMNS_XLS = 256;

	/**
	 * Maximum columns in .xlsx format
	 */
	public static final int MAX_COLUMNS_XLSX = 16384;

	// ==================== Default Values ====================

	/**
	 * Default validation start row (skip header)
	 */
	public static final int DEFAULT_VALIDATION_START_ROW = 1;

	/**
	 * Default validation end row
	 */
	public static final int DEFAULT_VALIDATION_END_ROW = 65535;

	/**
	 * Default header row index (0-based)
	 */
	public static final int DEFAULT_HEADER_ROW_INDEX = 0;

	/**
	 * Default first data row index (0-based, after header)
	 */
	public static final int DEFAULT_FIRST_DATA_ROW_INDEX = 1;

	/**
	 * Default column width in characters
	 */
	public static final int DEFAULT_COLUMN_WIDTH = 15;

	/**
	 * Default zoom scale (percentage)
	 */
	public static final int DEFAULT_ZOOM_SCALE = 100;

	/**
	 * Default batch size for processing
	 */
	public static final int DEFAULT_BATCH_SIZE = 500;

	// ==================== Date Formats ====================

	/**
	 * Date format: yyyy-MM-dd
	 */
	public static final String DATE_FORMAT_YMD = "yyyy-MM-dd";

	/**
	 * Date format: yyyy/MM/dd
	 */
	public static final String DATE_FORMAT_YMD_SLASH = "yyyy/MM/dd";

	/**
	 * DateTime format: yyyy-MM-dd HH:mm:ss
	 */
	public static final String DATETIME_FORMAT_YMDHMS = "yyyy-MM-dd HH:mm:ss";

	/**
	 * DateTime format: yyyy/MM/dd HH:mm:ss
	 */
	public static final String DATETIME_FORMAT_YMDHMS_SLASH = "yyyy/MM/dd HH:mm:ss";

	/**
	 * Time format: HH:mm:ss
	 */
	public static final String TIME_FORMAT_HMS = "HH:mm:ss";

	/**
	 * Default date format for validation
	 */
	public static final String DEFAULT_DATE_FORMAT = DATE_FORMAT_YMD;

	/**
	 * Minimum date for validation (Excel epoch)
	 */
	public static final String MIN_DATE = "1900-01-01";

	/**
	 * Maximum date for validation
	 */
	public static final String MAX_DATE = "2099-12-31";

	// ==================== Number Formats ====================

	/**
	 * Integer format
	 */
	public static final String NUMBER_FORMAT_INTEGER = "#,##0";

	/**
	 * Decimal format (2 decimal places)
	 */
	public static final String NUMBER_FORMAT_DECIMAL_2 = "#,##0.00";

	/**
	 * Percentage format (2 decimal places)
	 */
	public static final String NUMBER_FORMAT_PERCENTAGE = "0.00%";

	/**
	 * Currency format (USD)
	 */
	public static final String NUMBER_FORMAT_CURRENCY_USD = "$#,##0.00";

	/**
	 * Currency format (CNY)
	 */
	public static final String NUMBER_FORMAT_CURRENCY_CNY = "¥#,##0.00";

	// ==================== Style Constants ====================

	/**
	 * Default font name
	 */
	public static final String DEFAULT_FONT_NAME = "Arial";

	/**
	 * Default font size
	 */
	public static final short DEFAULT_FONT_SIZE = 11;

	/**
	 * Header font size
	 */
	public static final short HEADER_FONT_SIZE = 12;

	/**
	 * Default row height in points
	 */
	public static final short DEFAULT_ROW_HEIGHT = 20;

	/**
	 * Header row height in points
	 */
	public static final short HEADER_ROW_HEIGHT = 25;

	// ==================== Color Constants (Hex) ====================

	/**
	 * Header background color (light blue)
	 */
	public static final String COLOR_HEADER_BG = "#4472C4";

	/**
	 * Header font color (white)
	 */
	public static final String COLOR_HEADER_FONT = "#FFFFFF";

	/**
	 * Success/Green color
	 */
	public static final String COLOR_SUCCESS = "#63BE7B";

	/**
	 * Warning/Yellow color
	 */
	public static final String COLOR_WARNING = "#FFEB84";

	/**
	 * Error/Red color
	 */
	public static final String COLOR_ERROR = "#F8696B";

	/**
	 * Info/Blue color
	 */
	public static final String COLOR_INFO = "#4472C4";

	// ==================== Cache Constants ====================

	/**
	 * Default cache maximum size
	 */
	public static final int CACHE_MAX_SIZE = 1000;

	/**
	 * Default cache expiration time in minutes
	 */
	public static final int CACHE_EXPIRE_MINUTES = 60;

	// ==================== Error Messages ====================

	/**
	 * Error: Data class cannot be null
	 */
	public static final String ERR_DATA_CLASS_NULL = "Data class cannot be null";

	/**
	 * Error: Field not found
	 */
	public static final String ERR_FIELD_NOT_FOUND = "Field not found: %s";

	/**
	 * Error: Field access denied
	 */
	public static final String ERR_FIELD_ACCESS_DENIED = "Cannot access field: %s";

	/**
	 * Error: Invalid row range
	 */
	public static final String ERR_INVALID_ROW_RANGE = "Invalid row range: firstRow=%d, lastRow=%d";

	/**
	 * Error: Invalid column index
	 */
	public static final String ERR_INVALID_COLUMN_INDEX = "Invalid column index: %d";

	/**
	 * Error: Validation options empty
	 */
	public static final String ERR_VALIDATION_OPTIONS_EMPTY = "Validation options cannot be empty for LIST type";

	/**
	 * Error: Formula empty
	 */
	public static final String ERR_FORMULA_EMPTY = "Formula cannot be empty for FORMULA type";

	// ==================== File Extensions ====================

	/**
	 * Excel 97-2003 file extension
	 */
	public static final String FILE_EXT_XLS = ".xls";

	/**
	 * Excel 2007+ file extension
	 */
	public static final String FILE_EXT_XLSX = ".xlsx";

	/**
	 * CSV file extension
	 */
	public static final String FILE_EXT_CSV = ".csv";

	// ==================== MIME Types ====================

	/**
	 * MIME type for .xls files
	 */
	public static final String MIME_TYPE_XLS = "application/vnd.ms-excel";

	/**
	 * MIME type for .xlsx files
	 */
	public static final String MIME_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	/**
	 * MIME type for CSV files
	 */
	public static final String MIME_TYPE_CSV = "text/csv";

	// ==================== Sheet Names ====================

	/**
	 * Default sheet name
	 */
	public static final String DEFAULT_SHEET_NAME = "Sheet1";

	/**
	 * Maximum sheet name length
	 */
	public static final int MAX_SHEET_NAME_LENGTH = 31;

	// ==================== Image Constants ====================

	/**
	 * Default image width in pixels
	 */
	public static final int DEFAULT_IMAGE_WIDTH = 100;

	/**
	 * Default image height in pixels
	 */
	public static final int DEFAULT_IMAGE_HEIGHT = 100;

	/**
	 * Maximum image width in pixels
	 */
	public static final int MAX_IMAGE_WIDTH = 1024;

	/**
	 * Maximum image height in pixels
	 */
	public static final int MAX_IMAGE_HEIGHT = 1024;

	// ==================== Comment Constants ====================

	/**
	 * Default comment box width (in columns)
	 */
	public static final int DEFAULT_COMMENT_WIDTH = 2;

	/**
	 * Default comment box height (in rows)
	 */
	public static final int DEFAULT_COMMENT_HEIGHT = 3;

	/**
	 * Default comment author
	 */
	public static final String DEFAULT_COMMENT_AUTHOR = "System";

}
