package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Sheet Style Annotation
 * <p>
 * Configures sheet-level features like freeze panes, auto-filter, column width, etc.
 * </p>
 *
 * <p>Usage Examples:</p>
 * <pre>
 * // Freeze first row (header)
 * &#64;ExcelSheetStyle(freezeRow = 1)
 * public class ProductDTO { ... }
 *
 * // Freeze first 2 columns
 * &#64;ExcelSheetStyle(freezeColumn = 2)
 * public class OrderDTO { ... }
 *
 * // Freeze first row and first 2 columns
 * &#64;ExcelSheetStyle(freezeRow = 1, freezeColumn = 2)
 * public class SalesDTO { ... }
 *
 * // Enable auto-filter on header row
 * &#64;ExcelSheetStyle(autoFilter = true)
 * public class UserDTO { ... }
 *
 * // Combine features
 * &#64;ExcelSheetStyle(
 *     freezeRow = 1,
 *     freezeColumn = 1,
 *     autoFilter = true,
 *     defaultColumnWidth = 15
 * )
 * public class DataDTO { ... }
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheetStyle {

	/**
	 * Number of rows to freeze (from top)
	 * <p>
	 * 0 = no freeze, 1 = freeze header row, 2 = freeze first 2 rows, etc.
	 * </p>
	 *
	 * @return Number of rows to freeze
	 */
	int freezeRow() default 0;

	/**
	 * Number of columns to freeze (from left)
	 * <p>
	 * 0 = no freeze, 1 = freeze first column, 2 = freeze first 2 columns, etc.
	 * </p>
	 *
	 * @return Number of columns to freeze
	 */
	int freezeColumn() default 0;

	/**
	 * Enable auto-filter on header row
	 * <p>
	 * Creates dropdown filter buttons on header cells
	 * </p>
	 *
	 * @return true to enable auto-filter, false otherwise
	 */
	boolean autoFilter() default false;

	/**
	 * Auto-filter start row (0-based)
	 * <p>
	 * Only used when autoFilter is true
	 * </p>
	 *
	 * @return Start row for auto-filter
	 */
	int autoFilterStartRow() default 0;

	/**
	 * Default column width for all columns
	 * <p>
	 * Unit: number of characters (Excel default is 8.43)
	 * Use -1 to keep default behavior
	 * </p>
	 *
	 * @return Default column width
	 */
	int defaultColumnWidth() default -1;

	/**
	 * Zoom scale (percentage)
	 * <p>
	 * Range: 10-400, default is 100
	 * </p>
	 *
	 * @return Zoom scale
	 */
	int zoomScale() default 100;

	/**
	 * Show grid lines
	 *
	 * @return true to show grid lines, false to hide
	 */
	boolean showGridLines() default true;

}
