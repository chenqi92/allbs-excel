package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Formula Annotation
 * <p>
 * Used to mark fields that should contain Excel formulas instead of values.
 * Supports variable replacement for dynamic formula generation.
 * </p>
 *
 * <p>Usage Examples:</p>
 * <pre>
 * // Simple formula with row placeholder
 * &#64;ExcelProperty("Total Price")
 * &#64;ExcelFormula("=C{row}*D{row}")  // {row} will be replaced with current row number
 * private BigDecimal totalPrice;
 *
 * // Formula with column references
 * &#64;ExcelProperty("Grand Total")
 * &#64;ExcelFormula("=SUM(E2:E{lastRow})")  // {lastRow} will be replaced with last data row
 * private BigDecimal grandTotal;
 *
 * // Formula with fixed cell references
 * &#64;ExcelProperty("Tax Amount")
 * &#64;ExcelFormula("=E{row}*$F$1")  // Mix of relative and absolute references
 * private BigDecimal taxAmount;
 *
 * // Conditional formula
 * &#64;ExcelProperty("Status")
 * &#64;ExcelFormula("=IF(D{row}&gt;100,\"High\",\"Low\")")
 * private String status;
 * </pre>
 *
 * <p>Supported Placeholders:</p>
 * <ul>
 *   <li>{row} - Current row number (data row, excluding header)</li>
 *   <li>{lastRow} - Last data row number</li>
 *   <li>{col} - Current column letter (A, B, C, ...)</li>
 *   <li>{A}, {B}, {C}, ... - Specific column letters</li>
 * </ul>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelFormula {

	/**
	 * Formula template
	 * <p>
	 * Supports placeholders: {row}, {lastRow}, {col}, {A}, {B}, {C}, etc.
	 * </p>
	 * <p>Examples:</p>
	 * <ul>
	 *   <li>="C{row}*D{row}" - Multiply values in columns C and D</li>
	 *   <li>="SUM(E2:E{lastRow})" - Sum all values in column E</li>
	 *   <li>="IF(D{row}&gt;0,\"Yes\",\"No\")" - Conditional formula</li>
	 * </ul>
	 *
	 * @return Formula template
	 */
	String value();

	/**
	 * Whether this formula field should be enabled
	 *
	 * @return true if enabled, false otherwise
	 */
	boolean enabled() default true;

	/**
	 * Whether to apply formula only to specific rows
	 * <p>
	 * If true, the formula will only be applied to rows specified by startRow and endRow.
	 * If false (default), the formula will be applied to all data rows.
	 * </p>
	 *
	 * @return true to limit application, false otherwise
	 */
	boolean limitToRange() default false;

	/**
	 * Start row for formula application (1-based, excluding header)
	 * <p>
	 * Only used when limitToRange is true.
	 * </p>
	 *
	 * @return Start row number
	 */
	int startRow() default 1;

	/**
	 * End row for formula application (1-based, excluding header)
	 * <p>
	 * Only used when limitToRange is true.
	 * Use -1 for last row.
	 * </p>
	 *
	 * @return End row number
	 */
	int endRow() default -1;

}
