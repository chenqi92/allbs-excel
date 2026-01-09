package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 数值计算校验注解（类级别）
 * <p>
 * 校验多个数值字段的计算结果是否等于目标字段
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;FieldsCalculation(
 *     fields = {"materialCost", "laborCost", "adminCost"},
 *     resultField = "totalCost",
 *     operator = CalculationOperator.SUM,
 *     message = "总费用必须等于各项费用之和"
 * )
 * public class ExpenseDTO {
 *     private BigDecimal materialCost;
 *     private BigDecimal laborCost;
 *     private BigDecimal adminCost;
 *     private BigDecimal totalCost;
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(FieldsCalculation.List.class)
@Documented
public @interface FieldsCalculation {

    /**
     * 参与计算的字段名列表
     *
     * @return 字段名数组
     */
    String[] fields();

    /**
     * 结果字段名
     *
     * @return 字段名
     */
    String resultField();

    /**
     * 计算操作符
     *
     * @return 操作符
     */
    CalculationOperator operator();

    /**
     * 允许的误差范围（用于浮点数比较）
     *
     * @return 误差范围
     */
    double tolerance() default 0.001;

    /**
     * 是否忽略空值（空值视为 0）
     *
     * @return 是否忽略空值
     */
    boolean ignoreNull() default true;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "计算结果校验失败";

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
        FieldsCalculation[] value();
    }
}
