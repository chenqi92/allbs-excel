package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * SpEL 表达式校验注解（类级别）
 * <p>
 * 使用 Spring Expression Language (SpEL) 进行复杂的跨字段校验
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 校验总金额 = 单价 × 数量
 * &#64;CrossFieldExpression(
 *     expression = "totalAmount == unitPrice * quantity",
 *     fields = {"unitPrice", "quantity", "totalAmount"},
 *     message = "总金额必须等于单价乘以数量"
 * )
 * public class OrderLineDTO {
 *     private BigDecimal unitPrice;
 *     private Integer quantity;
 *     private BigDecimal totalAmount;
 * }
 *
 * // 使用根对象访问字段
 * &#64;CrossFieldExpression(
 *     expression = "#root.endDate == null or #root.startDate \u003c= #root.endDate",
 *     message = "开始日期不能大于结束日期"
 * )
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(CrossFieldExpression.List.class)
@Documented
public @interface CrossFieldExpression {

    /**
     * SpEL 表达式
     * <p>
     * 表达式应返回 boolean 类型，true 表示校验通过<br>
     * 可使用 #root 访问当前对象，如 #root.fieldName<br>
     * 也可直接使用字段名（需在 fields 中声明）
     * </p>
     *
     * @return SpEL 表达式
     */
    String expression();

    /**
     * 表达式中使用的字段名列表
     * <p>
     * 声明后可在表达式中直接使用字段名，无需 #root 前缀
     * </p>
     *
     * @return 字段名数组
     */
    String[] fields() default {};

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "表达式校验失败";

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
        CrossFieldExpression[] value();
    }
}
