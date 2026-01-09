package cn.allbs.excel.annotation.cross;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 条件定义注解
 * <p>
 * 用于在 {@link AllowedValuesIf} 等注解中定义多个依赖条件
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target({})
@Retention(RetentionPolicy.RUNTIME)
public @interface Condition {

    /**
     * 依赖的字段名
     *
     * @return 字段名
     */
    String field();

    /**
     * 依赖字段应该具有的值
     * <p>
     * 当依赖字段的值等于此值时，条件成立
     * </p>
     *
     * @return 期望的值
     */
    String value();

    /**
     * 是否为非空条件
     * <p>
     * 当设置为 true 时，只要依赖字段非空即认为条件成立，忽略 value 属性
     * </p>
     *
     * @return 是否为非空条件
     */
    boolean notEmpty() default false;

    /**
     * 是否为空条件
     * <p>
     * 当设置为 true 时，只有依赖字段为空时条件才成立
     * </p>
     *
     * @return 是否为空条件
     */
    boolean isEmpty() default false;
}
