package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 单元格合并注解
 * <p>
 * 用于标记需要进行单元格合并的字段，相同值的单元格会自动合并
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("部门")
 * &#64;ExcelMerge
 * private String department;
 *
 * &#64;ExcelProperty("姓名")
 * &#64;ExcelMerge(dependOn = "department")  // 依赖部门列，只有部门相同时才合并姓名
 * private String name;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelMerge {

    /**
     * 是否启用合并
     * <p>
     * 默认为 true，设置为 false 可以临时禁用合并
     * </p>
     *
     * @return boolean
     */
    boolean enabled() default true;

    /**
     * 依赖的字段名称
     * <p>
     * 当指定了依赖字段时，只有依赖字段的值相同时，当前字段才会合并
     * </p>
     * <p>
     * 例如：部门-姓名-职位 三列，姓名依赖部门，职位依赖姓名
     * </p>
     * <p>
     * 示例：
     * <pre>
     * &#64;ExcelProperty("部门")
     * &#64;ExcelMerge
     * private String department;
     *
     * &#64;ExcelProperty("姓名")
     * &#64;ExcelMerge(dependOn = "department")
     * private String name;
     *
     * &#64;ExcelProperty("职位")
     * &#64;ExcelMerge(dependOn = "name")
     * private String position;
     * </pre>
     * </p>
     *
     * @return 依赖的字段名称数组
     */
    String[] dependOn() default {};
}

