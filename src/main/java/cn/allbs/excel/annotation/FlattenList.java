package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * List 实体自动展开注解
 * <p>
 * 用于将 List 中的实体对象展开为多行，并自动合并非 List 字段的单元格
 * </p>
 *
 * <p>单个 List 展开示例：</p>
 * <pre>
 * public class Order {
 *     &#64;ExcelProperty("订单号")
 *     private String orderNo;
 *
 *     &#64;FlattenList  // 自动展开订单明细
 *     private List&lt;OrderItem&gt; items;
 * }
 *
 * public class OrderItem {
 *     &#64;ExcelProperty("商品名称")
 *     private String productName;
 *
 *     &#64;ExcelProperty("数量")
 *     private Integer quantity;
 * }
 *
 * // 导出结果：
 * // 订单号     | 商品名称 | 数量
 * // ORDER001  | 商品A    | 2
 * // (合并)    | 商品B    | 1
 * // (合并)    | 商品C    | 3
 * </pre>
 *
 * <p>多个 List 展开示例：</p>
 * <pre>
 * public class Student {
 *     &#64;ExcelProperty("姓名")
 *     private String name;
 *
 *     &#64;FlattenList
 *     private List&lt;Course&gt; courses;  // 3个课程
 *
 *     &#64;FlattenList
 *     private List&lt;Award&gt; awards;    // 2个奖项
 * }
 *
 * // 导出结果（取最长策略）：
 * // 姓名   | 课程   | 成绩 | 奖项   | 时间
 * // 张三   | 数学   | 95   | 一等奖 | 2023
 * // (合并) | 英语   | 88   | 二等奖 | 2024
 * // (合并) | 物理   | 92   | (空)   | (空)
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface FlattenList {

    /**
     * 字段名前缀
     * <p>
     * 展开后的字段名会添加此前缀，用于避免字段名冲突
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
     * 当 List 为空或 null 时的处理策略
     * <p>
     * KEEP_ROW: 保留一行，显示空值（默认）
     * SKIP_ROW: 跳过整行
     * </p>
     *
     * @return 空值处理策略
     */
    EmptyListStrategy emptyStrategy() default EmptyListStrategy.KEEP_ROW;

    /**
     * 多个 List 字段的合并策略
     * <p>
     * MAX_LENGTH: 按最长的 List 展开，短的补空（默认）
     * MIN_LENGTH: 按最短的 List 展开，长的截断
     * CARTESIAN: 笛卡尔积（慎用，数据量会指数增长）
     * </p>
     *
     * @return 合并策略
     */
    MultiListStrategy multiListStrategy() default MultiListStrategy.MAX_LENGTH;

    /**
     * 字段顺序
     * <p>
     * 控制 List 展开字段在表头中的位置
     * 数值越小越靠前
     * </p>
     *
     * @return 顺序值
     */
    int order() default 0;

    /**
     * 最大展开行数
     * <p>
     * 限制 List 展开的最大行数，防止数据过多
     * 0 表示不限制（默认）
     * </p>
     *
     * @return 最大行数
     */
    int maxRows() default 0;

    /**
     * 是否启用合并单元格
     * <p>
     * true: 合并非 List 字段的单元格（默认）
     * false: 不合并，每行都重复显示
     * </p>
     *
     * @return 是否合并
     */
    boolean mergeCell() default true;

    /**
     * 空 List 处理策略
     */
    enum EmptyListStrategy {
        /**
         * 保留一行，显示空值
         */
        KEEP_ROW,

        /**
         * 跳过整行
         */
        SKIP_ROW
    }

    /**
     * 多 List 合并策略
     */
    enum MultiListStrategy {
        /**
         * 按最长的 List 展开
         */
        MAX_LENGTH,

        /**
         * 按最短的 List 展开
         */
        MIN_LENGTH,

        /**
         * 笛卡尔积（所有组合）
         */
        CARTESIAN
    }
}
