package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 嵌套对象自动展开注解
 * <p>
 * 用于自动展开嵌套对象中所有标注了 @ExcelProperty 的字段
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 嵌套对象类
 * public class Department {
 *     &#64;ExcelProperty("部门编码")
 *     private String code;
 *
 *     &#64;ExcelProperty("部门名称")
 *     private String name;
 *
 *     private String internalId; // 无注解，不会被导出
 * }
 *
 * // 使用自动展开
 * public class User {
 *     &#64;ExcelProperty("姓名")
 *     private String name;
 *
 *     &#64;FlattenProperty  // 自动展开 Department 中所有有 @ExcelProperty 的字段
 *     private Department dept;
 * }
 *
 * // 导出结果包含：姓名、部门编码、部门名称
 * </pre>
 *
 * <p>字段名前缀示例：</p>
 * <pre>
 * // 使用前缀避免字段名冲突
 * &#64;FlattenProperty(prefix = "部门-")
 * private Department dept;
 *
 * // 导出结果：姓名、部门-部门编码、部门-部门名称
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface FlattenProperty {

    /**
     * 字段名前缀
     * <p>
     * 展开后的字段名会添加此前缀，用于避免字段名冲突
     * 例如：prefix = "部门-"，则 "部门编码" 会变成 "部门-部门编码"
     * </p>
     *
     * @return 前缀字符串
     */
    String prefix() default "";

    /**
     * 字段名后缀
     * <p>
     * 展开后的字段名会添加此后缀
     * </p>
     *
     * @return 后缀字符串
     */
    String suffix() default "";

    /**
     * 是否递归展开
     * <p>
     * 如果嵌套对象的字段也是对象且有 @FlattenProperty 注解，是否继续展开
     * 默认 false，避免无限递归
     * </p>
     *
     * @return 是否递归展开
     */
    boolean recursive() default false;

    /**
     * 最大递归深度
     * <p>
     * 当 recursive = true 时生效，限制递归深度
     * 默认 3 层
     * </p>
     *
     * @return 最大递归深度
     */
    int maxDepth() default 3;

    /**
     * 字段顺序偏移量
     * <p>
     * 用于控制展开字段在表头中的位置
     * 正数：向后移动
     * 负数：向前移动
     * 0：保持原位置（默认）
     * </p>
     *
     * @return 顺序偏移量
     */
    int orderOffset() default 0;

    /**
     * 当嵌套对象为 null 时的处理策略
     * <p>
     * EMPTY_STRING: 所有展开的字段都填充空字符串（默认）
     * SKIP: 跳过不填充
     * CUSTOM: 使用自定义值（通过 nullValue 指定）
     * </p>
     *
     * @return null 值处理策略
     */
    NullStrategy nullStrategy() default NullStrategy.EMPTY_STRING;

    /**
     * 自定义 null 值
     * <p>
     * 当 nullStrategy = CUSTOM 时使用
     * </p>
     *
     * @return 自定义 null 值
     */
    String nullValue() default "";

    /**
     * Null 值处理策略
     */
    enum NullStrategy {
        /**
         * 空字符串
         */
        EMPTY_STRING,

        /**
         * 跳过不填充
         */
        SKIP,

        /**
         * 自定义值
         */
        CUSTOM
    }
}
