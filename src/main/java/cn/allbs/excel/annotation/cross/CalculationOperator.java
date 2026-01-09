package cn.allbs.excel.annotation.cross;

/**
 * 计算操作符枚举
 * <p>
 * 用于 {@link FieldsCalculation} 注解中指定数值计算类型
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public enum CalculationOperator {

    /**
     * 求和
     */
    SUM("求和"),

    /**
     * 乘积
     */
    MULTIPLY("乘积"),

    /**
     * 平均值
     */
    AVERAGE("平均值"),

    /**
     * 最大值
     */
    MAX("最大值"),

    /**
     * 最小值
     */
    MIN("最小值"),

    /**
     * 差值（第一个减去其他所有）
     */
    SUBTRACT("差值");

    private final String description;

    CalculationOperator(String description) {
        this.description = description;
    }

    public String getDescription() {
        return description;
    }
}
