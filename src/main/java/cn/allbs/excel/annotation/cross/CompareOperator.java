package cn.allbs.excel.annotation.cross;

/**
 * 比较操作符枚举
 * <p>
 * 用于 {@link FieldsCompare} 注解中指定字段间的比较关系
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public enum CompareOperator {

    /**
     * 等于
     */
    EQUAL("==", "等于"),

    /**
     * 不等于
     */
    NOT_EQUAL("!=", "不等于"),

    /**
     * 小于
     */
    LESS_THAN("<", "小于"),

    /**
     * 小于或等于
     */
    LESS_THAN_OR_EQUAL("<=", "小于或等于"),

    /**
     * 大于
     */
    GREATER_THAN(">", "大于"),

    /**
     * 大于或等于
     */
    GREATER_THAN_OR_EQUAL(">=", "大于或等于");

    private final String symbol;
    private final String description;

    CompareOperator(String symbol, String description) {
        this.symbol = symbol;
        this.description = description;
    }

    public String getSymbol() {
        return symbol;
    }

    public String getDescription() {
        return description;
    }
}
