package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 动态表头注解
 * <p>
 * 用于标记需要动态生成表头的字段（通常是 Map 或自定义属性集合）
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 示例1：从数据中自动提取表头
 * &#64;DynamicHeaders(
 *     strategy = DynamicHeaderStrategy.FROM_DATA,
 *     headerPrefix = "属性-"
 * )
 * private Map&lt;String, Object&gt; dynamicFields;
 *
 * // 示例2：使用预定义表头
 * &#64;DynamicHeaders(
 *     strategy = DynamicHeaderStrategy.FROM_CONFIG,
 *     headers = {"扩展字段1", "扩展字段2", "扩展字段3"}
 * )
 * private Map&lt;String, Object&gt; extFields;
 *
 * // 示例3：混合模式
 * &#64;DynamicHeaders(
 *     strategy = DynamicHeaderStrategy.MIXED,
 *     headers = {"必需字段1", "必需字段2"},
 *     headerPrefix = "自定义-"
 * )
 * private Map&lt;String, Object&gt; properties;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface DynamicHeaders {

	/**
	 * 动态表头生成策略
	 * <p>
	 * 默认为 FROM_DATA，即从数据中自动提取表头
	 * </p>
	 *
	 * @return 生成策略
	 */
	DynamicHeaderStrategy strategy() default DynamicHeaderStrategy.FROM_DATA;

	/**
	 * 预定义表头列表
	 * <p>
	 * 当 strategy 为 FROM_CONFIG 或 MIXED 时使用
	 * </p>
	 *
	 * @return 表头数组
	 */
	String[] headers() default {};

	/**
	 * 表头前缀
	 * <p>
	 * 为动态生成的表头添加统一前缀，默认为空
	 * </p>
	 * <p>例如：headerPrefix = "自定义-"，则表头为 "自定义-字段名"</p>
	 *
	 * @return 表头前缀
	 */
	String headerPrefix() default "";

	/**
	 * 表头后缀
	 * <p>
	 * 为动态生成的表头添加统一后缀，默认为空
	 * </p>
	 * <p>例如：headerSuffix = "-属性"，则表头为 "字段名-属性"</p>
	 *
	 * @return 表头后缀
	 */
	String headerSuffix() default "";

	/**
	 * 源字段名
	 * <p>
	 * 指定从哪个字段中提取动态表头，默认为当前字段
	 * </p>
	 * <p>当字段本身不是 Map 而是包含 Map 的对象时使用</p>
	 *
	 * @return 源字段名
	 */
	String sourceField() default "";

	/**
	 * 键的排序方式
	 * <p>
	 * 对动态生成的表头进行排序
	 * </p>
	 * <ul>
	 *   <li>NONE: 不排序，保持插入顺序或数据顺序</li>
	 *   <li>ASC: 升序排序</li>
	 *   <li>DESC: 降序排序</li>
	 * </ul>
	 *
	 * @return 排序方式
	 */
	SortOrder order() default SortOrder.NONE;

	/**
	 * 最大列数限制
	 * <p>
	 * 限制动态生成的列数，防止数据过多导致性能问题
	 * </p>
	 * <p>默认为 -1，表示不限制</p>
	 *
	 * @return 最大列数
	 */
	int maxColumns() default -1;

	/**
	 * 是否启用
	 * <p>
	 * 默认启用，可通过设置为 false 临时禁用动态表头
	 * </p>
	 *
	 * @return 是否启用
	 */
	boolean enabled() default true;

	/**
	 * 排序方式枚举
	 */
	enum SortOrder {
		/**
		 * 不排序
		 */
		NONE,
		/**
		 * 升序
		 */
		ASC,
		/**
		 * 降序
		 */
		DESC
	}

}
