package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Watermark Annotation
 * <p>
 * Adds watermark to Excel sheets
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target({ElementType.METHOD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelWatermark {

	/**
	 * Watermark text
	 * <p>
	 * Supports SpEL expressions and placeholders:
	 * - ${user.name}: Current username
	 * - ${date}: Current date
	 * - ${time}: Current time
	 * - ${datetime}: Current datetime
	 * </p>
	 * <p>
	 * Example: "CONFIDENTIAL - ${user.name} - ${datetime}"
	 * </p>
	 */
	String text();

	/**
	 * Whether watermark is enabled
	 */
	boolean enabled() default true;

	/**
	 * Font name
	 * <p>
	 * Default: Arial
	 * </p>
	 */
	String fontName() default "Arial";

	/**
	 * Font size
	 * <p>
	 * Default: 48
	 * </p>
	 */
	int fontSize() default 48;

	/**
	 * Font color (RGB hex format)
	 * <p>
	 * Default: #D3D3D3 (light gray)
	 * Example: #FF0000 (red), #00FF00 (green), #0000FF (blue)
	 * </p>
	 */
	String color() default "#D3D3D3";

	/**
	 * Rotation angle (degrees)
	 * <p>
	 * Default: -45 (diagonal from bottom-left to top-right)
	 * Range: -90 to 90
	 * </p>
	 */
	int rotation() default -45;

	/**
	 * Opacity (transparency)
	 * <p>
	 * Default: 0.3 (30% opacity)
	 * Range: 0.0 (fully transparent) to 1.0 (fully opaque)
	 * </p>
	 */
	double opacity() default 0.3;

	/**
	 * Horizontal spacing between watermarks
	 * <p>
	 * Default: 200 (pixels)
	 * </p>
	 */
	int horizontalSpacing() default 200;

	/**
	 * Vertical spacing between watermarks
	 * <p>
	 * Default: 150 (pixels)
	 * </p>
	 */
	int verticalSpacing() default 150;

	/**
	 * Starting row index for watermark
	 * <p>
	 * Default: 0 (start from first row)
	 * </p>
	 */
	int startRow() default 0;

	/**
	 * Starting column index for watermark
	 * <p>
	 * Default: 0 (start from first column)
	 * </p>
	 */
	int startColumn() default 0;
}
