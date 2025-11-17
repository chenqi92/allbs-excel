package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 条件样式注解
 * <p>
 * 根据单元格值自动应用不同的样式
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("状态")
 * &#64;ConditionalStyle(conditions = {
 *     &#64;Condition(value = "已完成", style = &#64;CellStyleDef(backgroundColor = "#00FF00")),
 *     &#64;Condition(value = "进行中", style = &#64;CellStyleDef(backgroundColor = "#FFFF00")),
 *     &#64;Condition(value = "已取消", style = &#64;CellStyleDef(backgroundColor = "#FF0000", fontColor = "#FFFFFF"))
 * })
 * private String status;
 *
 * &#64;ExcelProperty("分数")
 * &#64;ConditionalStyle(conditions = {
 *     &#64;Condition(value = "&gt;=90", style = &#64;CellStyleDef(backgroundColor = "#00FF00")),
 *     &#64;Condition(value = "&gt;=60", style = &#64;CellStyleDef(backgroundColor = "#FFFF00")),
 *     &#64;Condition(value = "&lt;60", style = &#64;CellStyleDef(backgroundColor = "#FF0000"))
 * })
 * private Integer score;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ConditionalStyle {

	/**
	 * 条件列表
	 * <p>
	 * 定义多个条件及其对应的样式
	 * </p>
	 *
	 * @return 条件数组
	 */
	Condition[] conditions();

	/**
	 * 是否启用
	 * <p>
	 * 默认启用，可通过设置为 false 临时禁用条件样式
	 * </p>
	 *
	 * @return 是否启用
	 */
	boolean enabled() default true;

}
