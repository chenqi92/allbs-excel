package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 值依赖校验注解
 * <p>
 * 当依赖字段满足指定条件时，被标注的字段只能填写指定的值
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 单条件：当支付方式为"银行转账"时，银行名称只能填指定银行
 * &#64;ExcelProperty("银行名称")
 * &#64;AllowedValuesIf(
 *     dependField = "paymentMethod",
 *     dependValue = "银行转账",
 *     allowedValues = {"工商银行", "建设银行", "农业银行"},
 *     message = "银行转账时，银行名称只能填指定银行"
 * )
 * private String bankName;
 *
 * // 多条件联合：当国家为"中国"且省份为"北京"时，区县只能填北京的区
 * &#64;ExcelProperty("区县")
 * &#64;AllowedValuesIf(
 *     conditions = {
 *         &#64;Condition(field = "country", value = "中国"),
 *         &#64;Condition(field = "province", value = "北京")
 *     },
 *     allowedValues = {"朝阳区", "海淀区", "东城区"},
 *     message = "中国北京地区只能填指定区县"
 * )
 * private String district;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(AllowedValuesIf.List.class)
@Documented
public @interface AllowedValuesIf {

    /**
     * 单条件模式：依赖的字段名
     * <p>
     * 与 {@link #conditions()} 二选一，conditions 优先级更高
     * </p>
     *
     * @return 字段名
     */
    String dependField() default "";

    /**
     * 单条件模式：依赖字段应该具有的值
     *
     * @return 期望的值
     */
    String dependValue() default "";

    /**
     * 多条件联合模式（AND 关系）
     * <p>
     * 当所有条件都满足时，校验才会触发
     * </p>
     *
     * @return 条件数组
     */
    Condition[] conditions() default {};

    /**
     * 允许的值列表
     *
     * @return 允许的值数组
     */
    String[] allowedValues() default {};

    /**
     * 允许的值的正则表达式
     * <p>
     * 当 allowedValues 为空时，使用此正则进行校验
     * </p>
     *
     * @return 正则表达式
     */
    String pattern() default "";

    /**
     * 数值最小值（用于数值类型字段）
     *
     * @return 最小值
     */
    double minValue() default Double.MIN_VALUE;

    /**
     * 数值最大值（用于数值类型字段）
     *
     * @return 最大值
     */
    double maxValue() default Double.MAX_VALUE;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "字段值不在允许的范围内";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};

    /**
     * 容器注解，支持在同一字段上使用多个 @AllowedValuesIf
     */
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface List {
        AllowedValuesIf[] value();
    }
}
