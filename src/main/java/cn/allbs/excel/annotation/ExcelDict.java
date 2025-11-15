package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 字典转换注解
 * <p>
 * 用于标记需要进行字典转换的字段，支持导入导出时的字典值转换
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("性别")
 * &#64;ExcelDict(dictType = "sys_user_sex")
 * private String sex;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDict {

    /**
     * 字典类型
     * <p>
     * 用于标识字典的类型，如：sys_user_sex、sys_user_status 等
     * </p>
     *
     * @return 字典类型
     */
    String dictType();

    /**
     * 分隔符
     * <p>
     * 当字段值包含多个字典值时使用的分隔符，默认为逗号
     * </p>
     * <p>例如：字段值为 "1,2,3"，导出时会转换为 "男,女,未知"</p>
     *
     * @return 分隔符
     */
    String separator() default ",";

    /**
     * 读取时的转换方向
     * <p>
     * true: 将字典标签转换为字典值（导入时使用）<br>
     * false: 将字典值转换为字典标签（导出时使用）
     * </p>
     *
     * @return 是否反向转换
     */
    boolean readConverterExp() default true;
}

