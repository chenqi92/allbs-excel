package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 字段值匹配校验注解
 * <p>
 * 被标注的字段值必须与另一个字段的值相同
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("密码")
 * private String password;
 *
 * &#64;ExcelProperty("确认密码")
 * &#64;FieldsMatch(field = "password", message = "两次输入的密码不一致")
 * private String confirmPassword;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface FieldsMatch {

    /**
     * 需要匹配的字段名
     *
     * @return 字段名
     */
    String field();

    /**
     * 是否忽略大小写
     *
     * @return 是否忽略大小写
     */
    boolean ignoreCase() default false;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "字段值不匹配";

    /**
     * 校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] groups() default {};
}
