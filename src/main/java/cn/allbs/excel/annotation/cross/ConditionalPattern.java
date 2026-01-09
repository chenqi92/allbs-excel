package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 条件正则校验注解
 * <p>
 * 根据依赖字段的值，应用不同的正则表达式进行校验
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("国家代码")
 * private String countryCode;
 *
 * &#64;ExcelProperty("手机号")
 * &#64;ConditionalPattern(
 *     dependField = "countryCode",
 *     patterns = {
 *         &#64;ConditionalPattern.PatternRule(value = "CN", pattern = "^1[3-9]\\d{9}$"),
 *         &#64;ConditionalPattern.PatternRule(value = "US", pattern = "^\\d{10}$"),
 *         &#64;ConditionalPattern.PatternRule(value = "JP", pattern = "^0[789]0\\d{8}$")
 *     },
 *     message = "手机号格式不正确"
 * )
 * private String phone;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ConditionalPattern {

    /**
     * 依赖的字段名
     *
     * @return 字段名
     */
    String dependField();

    /**
     * 正则规则列表
     *
     * @return 规则数组
     */
    PatternRule[] patterns();

    /**
     * 默认正则表达式
     * <p>
     * 当依赖字段的值不匹配任何 patterns 规则时使用
     * </p>
     *
     * @return 默认正则
     */
    String defaultPattern() default "";

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "格式校验失败";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};

    /**
     * 正则规则定义
     */
    @interface PatternRule {
        /**
         * 依赖字段的值
         */
        String value();

        /**
         * 对应的正则表达式
         */
        String pattern();
    }
}
