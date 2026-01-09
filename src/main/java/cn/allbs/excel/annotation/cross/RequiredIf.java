package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 条件必填校验注解
 * <p>
 * 当指定的依赖字段有值时，被标注的字段必须填写
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 当填写了收货地址时，收货人必填
 * &#64;ExcelProperty("收货地址")
 * private String address;
 *
 * &#64;ExcelProperty("收货人")
 * &#64;RequiredIf(field = "address", message = "填写了收货地址时，收货人必填")
 * private String receiver;
 *
 * // 当支付方式为"银行转账"时，银行名称必填
 * &#64;ExcelProperty("银行名称")
 * &#64;RequiredIf(field = "paymentMethod", hasValue = "银行转账", message = "银行转账时必须填写银行名称")
 * private String bankName;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(RequiredIf.List.class)
@Documented
public @interface RequiredIf {

    /**
     * 依赖的字段名
     *
     * @return 字段名
     */
    String field();

    /**
     * 依赖字段应该具有的值
     * <p>
     * 为空字符串时，表示只要依赖字段非空即触发校验
     * </p>
     *
     * @return 期望的值
     */
    String hasValue() default "";

    /**
     * 多条件联合（AND 关系）
     * <p>
     * 当需要多个字段同时满足条件时使用
     * </p>
     *
     * @return 条件数组
     */
    Condition[] conditions() default {};

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "条件必填校验失败";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};

    /**
     * 容器注解，支持在同一字段上使用多个 @RequiredIf
     */
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface List {
        RequiredIf[] value();
    }
}
