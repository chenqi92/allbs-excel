package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 互斥校验注解
 * <p>
 * 被标注的字段与指定字段之间只能填写一个。
 * </p>
 * <p>
 * <b>注意：</b>此注解是单向校验——仅在被标注字段非空时才检查互斥字段。
 * 如需双向互斥，需在两个字段上各标注一次，参见下方示例。
 * </p>
 *
 * <p>
 * 使用示例：
 * </p>
 * 
 * <pre>
 * // 手机号和座机号只能填一个
 * &#64;ExcelProperty("手机号")
 * &#64;MutualExclusive(fields = { "telephone" }, message = "手机号和座机号只能填一个")
 * private String mobile;
 *
 * &#64;ExcelProperty("座机号")
 * &#64;MutualExclusive(fields = { "mobile" }, message = "手机号和座机号只能填一个")
 * private String telephone;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface MutualExclusive {

    /**
     * 互斥的字段名列表
     *
     * @return 字段名数组
     */
    String[] fields();

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "这些字段只能填写一个";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};
}
