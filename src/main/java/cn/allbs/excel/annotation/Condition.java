package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 条件注解
 * <p>
 * 用于定义单元格样式的应用条件
 * </p>
 * <p>使用示例：</p>
 * <pre>
 * &#64;Condition(value = "已完成", style = &#64;CellStyleDef(backgroundColor = "#00FF00"))
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target({})
@Retention(RetentionPolicy.RUNTIME)
public @interface Condition {

	/**
	 * 条件值
	 * <p>
	 * 支持以下几种格式：
	 * </p>
	 * <ul>
	 *   <li>精确匹配：直接写值，如 "已完成"</li>
	 *   <li>大于：使用 &gt; 前缀，如 "&gt;100"</li>
	 *   <li>小于：使用 &lt; 前缀，如 "&lt;50"</li>
	 *   <li>大于等于：使用 &gt;= 前缀，如 "&gt;=60"</li>
	 *   <li>小于等于：使用 &lt;= 前缀，如 "&lt;=90"</li>
	 *   <li>区间：使用 [] 或 () 表示，如 "[60,90]" 表示 60到90（含边界）</li>
	 *   <li>正则表达式：使用 regex: 前缀，如 "regex:^[A-Z].*"</li>
	 *   <li>SpEL表达式：使用 spel: 前缀，如 "spel:#value > 100 and #value < 200"</li>
	 * </ul>
	 *
	 * @return 条件值
	 */
	String value();

	/**
	 * 应用的样式
	 *
	 * @return 单元格样式定义
	 */
	CellStyleDef style();

	/**
	 * 条件优先级
	 * <p>
	 * 当多个条件同时满足时，优先级高的样式会被应用
	 * 数值越小，优先级越高，默认为 0
	 * </p>
	 *
	 * @return 优先级
	 */
	int priority() default 0;

}
