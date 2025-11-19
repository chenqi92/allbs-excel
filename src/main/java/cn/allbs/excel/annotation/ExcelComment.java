package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Comment Annotation
 * <p>
 * Adds comments/notes to Excel cells.
 * Comments appear as small red triangles in cell corners and show text on hover.
 * </p>
 *
 * <p>Usage Examples:</p>
 * <pre>
 * // Simple comment on header
 * &#64;ExcelProperty("Email")
 * &#64;ExcelComment(
 *     headerComment = "Must be a valid email format",
 *     author = "System"
 * )
 * private String email;
 *
 * // Dynamic comment based on cell value
 * &#64;ExcelProperty("Status")
 * &#64;ExcelComment(
 *     headerComment = "Order status code",
 *     dataComment = "Status: {value}",
 *     author = "Admin",
 *     width = 3,
 *     height = 2
 * )
 * private String status;
 *
 * // Comment with custom size and visibility
 * &#64;ExcelProperty("Notes")
 * &#64;ExcelComment(
 *     headerComment = "Internal notes - not visible to customer",
 *     visible = true,
 *     width = 5,
 *     height = 3
 * )
 * private String notes;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelComment {

	/**
	 * Comment text for header cell
	 * <p>
	 * This comment will be added to the column header.
	 * Leave empty to skip header comment.
	 * </p>
	 *
	 * @return Header comment text
	 */
	String headerComment() default "";

	/**
	 * Comment text template for data cells
	 * <p>
	 * Supports placeholders:
	 * - {value} - Current cell value
	 * - {row} - Current row number
	 * - {col} - Current column letter
	 * </p>
	 * <p>
	 * Leave empty to skip data cell comments.
	 * </p>
	 *
	 * @return Data comment template
	 */
	String dataComment() default "";

	/**
	 * Comment author name
	 *
	 * @return Author name
	 */
	String author() default "System";

	/**
	 * Comment box width (in columns)
	 * <p>
	 * Default is 2 columns wide
	 * </p>
	 *
	 * @return Width in columns
	 */
	int width() default 2;

	/**
	 * Comment box height (in rows)
	 * <p>
	 * Default is 3 rows high
	 * </p>
	 *
	 * @return Height in rows
	 */
	int height() default 3;

	/**
	 * Whether comment is always visible
	 * <p>
	 * If false (default), comment only shows on hover.
	 * If true, comment is always visible.
	 * </p>
	 *
	 * @return true for always visible, false for hover only
	 */
	boolean visible() default false;

	/**
	 * Apply comment to all data rows
	 * <p>
	 * If true, dataComment will be applied to all rows.
	 * If false, only specific rows defined by startRow/endRow.
	 * </p>
	 *
	 * @return true to apply to all rows
	 */
	boolean applyToAll() default false;

	/**
	 * Start row for data comments (1-based, excluding header)
	 * Only used when applyToAll is false
	 *
	 * @return Start row number
	 */
	int startRow() default 1;

	/**
	 * End row for data comments (1-based, excluding header)
	 * Only used when applyToAll is false
	 * Use -1 for last row
	 *
	 * @return End row number
	 */
	int endRow() default -1;

}
