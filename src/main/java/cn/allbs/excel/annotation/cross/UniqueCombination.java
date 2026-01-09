package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 跨行唯一性校验注解（类级别）
 * <p>
 * 校验指定字段组合在整个 Excel 中是否唯一
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;UniqueCombination(
 *     fields = {"employeeId", "attendanceDate"},
 *     message = "同一员工同一天不能有重复考勤记录"
 * )
 * public class AttendanceDTO {
 *     private String employeeId;
 *     private LocalDate attendanceDate;
 *     private String attendanceType;
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(UniqueCombination.List.class)
@Documented
public @interface UniqueCombination {

    /**
     * 参与唯一性校验的字段名列表
     *
     * @return 字段名数组
     */
    String[] fields();

    /**
     * 是否忽略大小写（用于字符串字段）
     *
     * @return 是否忽略大小写
     */
    boolean ignoreCase() default false;

    /**
     * 是否忽略空值组合
     * <p>
     * 当设置为 true 时，如果组合中任一字段为空，则不进行唯一性校验
     * </p>
     *
     * @return 是否忽略空值
     */
    boolean ignoreNull() default false;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "数据重复";

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
        UniqueCombination[] value();
    }
}
