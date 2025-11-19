package cn.allbs.excel.annotation;

import java.lang.annotation.*;

/**
 * Conditional Format Annotation
 * <p>
 * Applies conditional formatting to cells based on their values.
 * Supports data bars, color scales, icon sets, and custom rules.
 * </p>
 *
 * <p>Usage Examples:</p>
 * <pre>
 * // Data bars (visual bars in cells)
 * &#64;ExcelProperty("Sales Amount")
 * &#64;ConditionalFormat(type = FormatType.DATA_BAR, color = "#4472C4")
 * private BigDecimal salesAmount;
 *
 * // Color scale (gradient based on value)
 * &#64;ExcelProperty("Score")
 * &#64;ConditionalFormat(
 *     type = FormatType.COLOR_SCALE,
 *     minColor = "#F8696B",  // Red for low values
 *     midColor = "#FFEB84",  // Yellow for medium values
 *     maxColor = "#63BE7B"   // Green for high values
 * )
 * private Integer score;
 *
 * // Icon set (arrows, traffic lights, etc.)
 * &#64;ExcelProperty("Status")
 * &#64;ConditionalFormat(
 *     type = FormatType.ICON_SET,
 *     iconSet = IconSetType.THREE_ARROWS
 * )
 * private Integer status;
 *
 * // Highlight cells above average
 * &#64;ExcelProperty("Revenue")
 * &#64;ConditionalFormat(
 *     type = FormatType.ABOVE_AVERAGE,
 *     fillColor = "#C6EFCE",
 *     fontColor = "#006100"
 * )
 * private BigDecimal revenue;
 *
 * // Highlight top 10%
 * &#64;ExcelProperty("Performance")
 * &#64;ConditionalFormat(
 *     type = FormatType.TOP_PERCENT,
 *     rank = 10,
 *     fillColor = "#FFC7CE",
 *     fontColor = "#9C0006"
 * )
 * private Double performance;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(ConditionalFormats.class)
public @interface ConditionalFormat {

	/**
	 * Conditional format type
	 */
	FormatType type();

	/**
	 * Apply to all rows in this column
	 * If false, only applies to specific row range
	 */
	boolean applyToAll() default true;

	/**
	 * Start row (1-based, excluding header)
	 * Only used when applyToAll is false
	 */
	int startRow() default 1;

	/**
	 * End row (1-based, excluding header)
	 * Only used when applyToAll is false
	 * Use -1 for last row
	 */
	int endRow() default -1;

	// ==================== Data Bar Settings ====================

	/**
	 * Color for data bar (hex format: #RRGGBB)
	 */
	String color() default "#4472C4";

	/**
	 * Show bar only (hide cell value)
	 */
	boolean showBarOnly() default false;

	// ==================== Color Scale Settings ====================

	/**
	 * Minimum value color (hex format: #RRGGBB)
	 */
	String minColor() default "#F8696B";

	/**
	 * Middle value color (hex format: #RRGGBB)
	 * Leave empty for 2-color scale
	 */
	String midColor() default "#FFEB84";

	/**
	 * Maximum value color (hex format: #RRGGBB)
	 */
	String maxColor() default "#63BE7B";

	// ==================== Icon Set Settings ====================

	/**
	 * Icon set type
	 */
	IconSetType iconSet() default IconSetType.THREE_ARROWS;

	/**
	 * Reverse icon order
	 */
	boolean reverseIcons() default false;

	/**
	 * Show icons only (hide cell value)
	 */
	boolean showIconOnly() default false;

	// ==================== Highlight Settings ====================

	/**
	 * Fill color (hex format: #RRGGBB)
	 */
	String fillColor() default "#FFEB9C";

	/**
	 * Font color (hex format: #RRGGBB)
	 */
	String fontColor() default "#9C6500";

	/**
	 * Rank/Percentage for top/bottom rules
	 */
	int rank() default 10;

	// ==================== Custom Formula ====================

	/**
	 * Custom formula for conditional formatting
	 * Example: "AND($B2>100, $C2<50)"
	 */
	String formula() default "";

	/**
	 * Format Type Enum
	 */
	enum FormatType {
		/**
		 * Data bars (horizontal bars in cells)
		 */
		DATA_BAR,

		/**
		 * Color scale (2 or 3 color gradient)
		 */
		COLOR_SCALE,

		/**
		 * Icon sets (arrows, traffic lights, etc.)
		 */
		ICON_SET,

		/**
		 * Highlight cells above average
		 */
		ABOVE_AVERAGE,

		/**
		 * Highlight cells below average
		 */
		BELOW_AVERAGE,

		/**
		 * Highlight top N items
		 */
		TOP_N,

		/**
		 * Highlight bottom N items
		 */
		BOTTOM_N,

		/**
		 * Highlight top N percent
		 */
		TOP_PERCENT,

		/**
		 * Highlight bottom N percent
		 */
		BOTTOM_PERCENT,

		/**
		 * Highlight duplicate values
		 */
		DUPLICATE_VALUES,

		/**
		 * Highlight unique values
		 */
		UNIQUE_VALUES,

		/**
		 * Custom formula-based formatting
		 */
		FORMULA
	}

	/**
	 * Icon Set Type Enum
	 */
	enum IconSetType {
		/**
		 * 3 arrows (up, right, down)
		 */
		THREE_ARROWS,

		/**
		 * 3 arrows (gray)
		 */
		THREE_ARROWS_GRAY,

		/**
		 * 3 flags
		 */
		THREE_FLAGS,

		/**
		 * 3 traffic lights (unrimmed)
		 */
		THREE_TRAFFIC_LIGHTS_1,

		/**
		 * 3 traffic lights (rimmed)
		 */
		THREE_TRAFFIC_LIGHTS_2,

		/**
		 * 3 signs
		 */
		THREE_SIGNS,

		/**
		 * 3 symbols (circled)
		 */
		THREE_SYMBOLS,

		/**
		 * 3 symbols (uncircled)
		 */
		THREE_SYMBOLS_2,

		/**
		 * 4 arrows (gray)
		 */
		FOUR_ARROWS_GRAY,

		/**
		 * 4 arrows
		 */
		FOUR_ARROWS,

		/**
		 * 4 rating
		 */
		FOUR_RATING,

		/**
		 * 4 red to black
		 */
		FOUR_RED_TO_BLACK,

		/**
		 * 4 traffic lights
		 */
		FOUR_TRAFFIC_LIGHTS,

		/**
		 * 5 arrows (gray)
		 */
		FIVE_ARROWS_GRAY,

		/**
		 * 5 arrows
		 */
		FIVE_ARROWS,

		/**
		 * 5 rating
		 */
		FIVE_RATING,

		/**
		 * 5 quarters
		 */
		FIVE_QUARTERS
	}

}
