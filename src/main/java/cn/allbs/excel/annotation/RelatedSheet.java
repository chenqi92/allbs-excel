package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 关联 Sheet 注解
 * <p>
 * 用于标记需要导出到关联 Sheet 的字段（通常是 List 或集合类型）
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * public class OrderDTO {
 *     &#64;ExcelProperty("订单号")
 *     private String orderNo;
 *
 *     &#64;RelatedSheet(
 *         sheetName = "订单明细",
 *         relationKey = "orderNo",
 *         createHyperlink = true
 *     )
 *     private List&lt;OrderItemDTO&gt; items;
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface RelatedSheet {

	/**
	 * 关联 Sheet 的名称
	 *
	 * @return Sheet 名称
	 */
	String sheetName();

	/**
	 * 关联键字段名（主表中的字段名）
	 * <p>
	 * 例如：主表为订单，关联表为订单明细，则 relationKey 为 "orderNo"
	 * </p>
	 *
	 * @return 关联键字段名
	 */
	String relationKey();

	/**
	 * 子表中对应的关联字段名（如果与主表字段名不同）
	 * <p>
	 * 例如：主表字段为 "orderNo"，子表字段为 "parentOrderNo"，则设置此值为 "parentOrderNo"<br>
	 * 如果不设置，默认使用 relationKey 的值
	 * </p>
	 *
	 * @return 子表关联字段名
	 */
	String childRelationKey() default "";

	/**
	 * 是否创建超链接
	 * <p>
	 * true: 在主表的关联列创建超链接，点击可跳转到子表<br>
	 * false: 不创建超链接
	 * </p>
	 *
	 * @return 是否创建超链接
	 */
	boolean createHyperlink() default true;

	/**
	 * 超链接显示文本
	 * <p>
	 * 如果为空，默认显示关联数据的数量，例如："查看明细(3)"
	 * </p>
	 *
	 * @return 超链接文本
	 */
	String hyperlinkText() default "";

	/**
	 * 子表数据的类型
	 * <p>
	 * 如果字段是泛型 List，需要指定具体的数据类型
	 * </p>
	 *
	 * @return 子表数据类型
	 */
	Class<?> dataType() default Object.class;

	/**
	 * 子表的排序字段
	 * <p>
	 * 指定子表数据的排序字段，默认不排序
	 * </p>
	 *
	 * @return 排序字段名
	 */
	String orderBy() default "";

	/**
	 * 是否启用
	 *
	 * @return 是否启用
	 */
	boolean enabled() default true;

}
