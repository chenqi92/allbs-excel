package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 嵌套字段展开注解
 * <p>
 * 用于标记需要展开的嵌套对象字段，支持多层嵌套路径表达式、集合索引访问、Map键访问等
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 示例1：单层嵌套
 * &#64;ExcelProperty("部门名称")
 * &#64;NestedProperty("dept.name")
 * private Department dept;
 *
 * // 示例2：多层嵌套
 * &#64;ExcelProperty("领导姓名")
 * &#64;NestedProperty("dept.leader.name")
 * private Department dept;
 *
 * // 示例3：集合索引访问
 * &#64;ExcelProperty("第一个部门")
 * &#64;NestedProperty("depts[0].name")
 * private List&lt;Department&gt; depts;
 *
 * // 示例4：集合字段拼接
 * &#64;ExcelProperty("所有部门")
 * &#64;NestedProperty(value = "depts[*].name", separator = ",")
 * private List&lt;Department&gt; depts;
 *
 * // 示例5：Map键访问
 * &#64;ExcelProperty("扩展属性")
 * &#64;NestedProperty("properties[city]")
 * private Map&lt;String, Object&gt; properties;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface NestedProperty {

    /**
     * 嵌套字段路径表达式
     * <p>
     * 支持以下语法：
     * <ul>
     *   <li>对象字段：dept.name</li>
     *   <li>集合索引：depts[0].name（访问第一个元素）</li>
     *   <li>集合全部：depts[*].name（访问所有元素并拼接）</li>
     *   <li>Map键访问：properties[key] 或 properties.key</li>
     *   <li>数组索引：tags[0]（访问数组元素）</li>
     * </ul>
     * </p>
     *
     * @return 嵌套字段路径
     */
    String value();

    /**
     * 当嵌套对象为 null 或字段值为 null 时的默认值
     * <p>
     * 默认为空字符串
     * </p>
     *
     * @return 空值时的默认值
     */
    String nullValue() default "";

    /**
     * 集合元素拼接分隔符
     * <p>
     * 当使用 [*] 访问集合所有元素时，使用此分隔符拼接结果
     * 默认为逗号
     * </p>
     *
     * @return 分隔符
     */
    String separator() default ",";

    /**
     * 是否忽略嵌套路径访问异常
     * <p>
     * true: 遇到异常时返回 nullValue<br>
     * false: 抛出异常
     * </p>
     *
     * @return 是否忽略异常
     */
    boolean ignoreException() default true;

    /**
     * 集合元素最大拼接数量
     * <p>
     * 当使用 [*] 访问集合所有元素时，最多拼接的元素数量
     * 0 表示不限制，默认为 0
     * </p>
     *
     * @return 最大拼接数量
     */
    int maxJoinSize() default 0;
}
