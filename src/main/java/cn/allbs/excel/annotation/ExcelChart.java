package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Chart Annotation
 * <p>
 * Creates charts in Excel sheets based on data
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target({ElementType.METHOD, ElementType.TYPE, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelChart {

	/**
	 * Chart title
	 */
	String title() default "";

	/**
	 * Whether chart is enabled
	 */
	boolean enabled() default true;

	/**
	 * Chart type
	 */
	ChartType type() default ChartType.LINE;

	/**
	 * X-axis field name
	 * <p>
	 * The field to use for X-axis (category axis)
	 * </p>
	 */
	String xAxisField() default "";

	/**
	 * Y-axis field names
	 * <p>
	 * The fields to use for Y-axis (value axis)
	 * Can specify multiple fields for multiple series
	 * </p>
	 */
	String[] yAxisFields() default {};

	/**
	 * Chart position - starting row
	 */
	int startRow() default 0;

	/**
	 * Chart position - starting column
	 */
	int startColumn() default 0;

	/**
	 * Chart position - ending row
	 */
	int endRow() default 15;

	/**
	 * Chart position - ending column
	 */
	int endColumn() default 10;

	/**
	 * X-axis title
	 */
	String xAxisTitle() default "";

	/**
	 * Y-axis title
	 */
	String yAxisTitle() default "";

	/**
	 * Whether to show legend
	 */
	boolean showLegend() default true;

	/**
	 * Legend position
	 */
	LegendPosition legendPosition() default LegendPosition.BOTTOM;

	/**
	 * Chart type enum
	 */
	enum ChartType {
		/** Line chart */
		LINE,
		/** Bar chart */
		BAR,
		/** Column chart */
		COLUMN,
		/** Pie chart */
		PIE,
		/** Area chart */
		AREA,
		/** Scatter chart */
		SCATTER
	}

	/**
	 * Legend position enum
	 */
	enum LegendPosition {
		/** Top */
		TOP,
		/** Bottom */
		BOTTOM,
		/** Left */
		LEFT,
		/** Right */
		RIGHT,
		/** Top right */
		TOP_RIGHT
	}
}
