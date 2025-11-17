package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Sheet 关系配置注解
 * <p>
 * 用于配置多个 Sheet 之间的关系（可用于类级别或方法级别）
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;SheetRelation(
 *     mainSheet = "订单",
 *     relatedSheets = {
 *         &#64;SheetRelation.Relation(
 *             sheetName = "订单明细",
 *             relationKey = "orderNo",
 *             createIndex = true
 *         )
 *     }
 * )
 * public class OrderExportDTO {
 *     // ...
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target({ ElementType.TYPE, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetRelation {

	/**
	 * 主 Sheet 名称
	 *
	 * @return 主 Sheet 名称
	 */
	String mainSheet() default "主表";

	/**
	 * 关联的 Sheet 配置
	 *
	 * @return 关联配置数组
	 */
	Relation[] relatedSheets() default {};

	/**
	 * 是否自动创建目录 Sheet
	 * <p>
	 * true: 创建一个目录 Sheet，包含所有 Sheet 的链接<br>
	 * false: 不创建目录
	 * </p>
	 *
	 * @return 是否创建目录
	 */
	boolean createIndex() default false;

	/**
	 * 目录 Sheet 名称
	 *
	 * @return 目录 Sheet 名称
	 */
	String indexSheetName() default "目录";

	/**
	 * 关联关系定义
	 */
	@Target({})
	@Retention(RetentionPolicy.RUNTIME)
	@interface Relation {

		/**
		 * Sheet 名称
		 *
		 * @return Sheet 名称
		 */
		String sheetName();

		/**
		 * 关联键（主表字段名）
		 *
		 * @return 关联键
		 */
		String relationKey();

		/**
		 * 子表关联键（如果与主表不同）
		 *
		 * @return 子表关联键
		 */
		String childRelationKey() default "";

		/**
		 * 是否创建超链接
		 *
		 * @return 是否创建超链接
		 */
		boolean createHyperlink() default true;

		/**
		 * 数据类型
		 *
		 * @return 数据类型
		 */
		Class<?> dataType() default Object.class;

	}

}
