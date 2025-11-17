package cn.allbs.excel.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 单元格样式定义注解
 * <p>
 * 用于定义 Excel 单元格的样式属性
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target({})
@Retention(RetentionPolicy.RUNTIME)
public @interface CellStyleDef {

	/**
	 * 前景色（十六进制格式，如 "#FF0000" 表示红色）
	 * <p>
	 * 默认为空，表示不设置前景色
	 * </p>
	 *
	 * @return 前景色
	 */
	String foregroundColor() default "";

	/**
	 * 背景色（十六进制格式，如 "#00FF00" 表示绿色）
	 * <p>
	 * 默认为空，表示不设置背景色
	 * </p>
	 *
	 * @return 背景色
	 */
	String backgroundColor() default "";

	/**
	 * 填充模式
	 * <p>
	 * 使用 POI 的 FillPatternType，默认为 SOLID_FOREGROUND
	 * </p>
	 *
	 * @return 填充模式
	 */
	short fillPattern() default 1; // FillPatternType.SOLID_FOREGROUND

	/**
	 * 字体颜色（十六进制格式，如 "#FFFFFF" 表示白色）
	 * <p>
	 * 默认为空，表示不设置字体颜色
	 * </p>
	 *
	 * @return 字体颜色
	 */
	String fontColor() default "";

	/**
	 * 字体是否加粗
	 * <p>
	 * true: 加粗<br>
	 * false: 不加粗（默认）
	 * </p>
	 *
	 * @return 是否加粗
	 */
	boolean bold() default false;

	/**
	 * 字体大小
	 * <p>
	 * 默认为 -1，表示使用默认字体大小
	 * </p>
	 *
	 * @return 字体大小
	 */
	short fontSize() default -1;

	/**
	 * 水平对齐方式
	 * <p>
	 * 使用 POI 的 HorizontalAlignment，默认为 -1（不设置）
	 * </p>
	 * <ul>
	 *   <li>1 = LEFT</li>
	 *   <li>2 = CENTER</li>
	 *   <li>3 = RIGHT</li>
	 * </ul>
	 *
	 * @return 水平对齐方式
	 */
	short horizontalAlignment() default -1;

	/**
	 * 垂直对齐方式
	 * <p>
	 * 使用 POI 的 VerticalAlignment，默认为 -1（不设置）
	 * </p>
	 * <ul>
	 *   <li>0 = TOP</li>
	 *   <li>1 = CENTER</li>
	 *   <li>2 = BOTTOM</li>
	 * </ul>
	 *
	 * @return 垂直对齐方式
	 */
	short verticalAlignment() default -1;

}
