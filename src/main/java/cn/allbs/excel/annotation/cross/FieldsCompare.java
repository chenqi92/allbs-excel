package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 字段比较校验注解（类级别）
 * <p>
 * 比较两个字段的值，支持数值和日期类型的比较
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;FieldsCompare(
 *     first = "startDate",
 *     second = "endDate",
 *     operator = CompareOperator.LESS_THAN_OR_EQUAL,
 *     message = "开始日期不能大于结束日期"
 * )
 * public class DateRangeDTO {
 *     private LocalDate startDate;
 *     private LocalDate endDate;
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(FieldsCompare.List.class)
@Documented
public @interface FieldsCompare {

    /**
     * 第一个字段名（比较的左侧）
     *
     * @return 字段名
     */
    String first();

    /**
     * 第二个字段名（比较的右侧）
     *
     * @return 字段名
     */
    String second();

    /**
     * 比较操作符
     *
     * @return 操作符
     */
    CompareOperator operator();

    /**
     * 是否允许两个字段都为空
     *
     * @return 是否允许都为空
     */
    boolean allowBothNull() default true;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "字段比较校验失败";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};

    /**
     * 容器注解
     */
    @Target(ElementType.TYPE)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface List {
        FieldsCompare[] value();
    }
}
