package cn.allbs.excel.annotation.cross;

import java.lang.annotation.*;

/**
 * 至少一个非空校验注解（类级别）
 * <p>
 * 指定的多个字段中至少有一个需要填写
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;AtLeastOne(
 *     fields = {"mobile", "email", "qq"},
 *     message = "手机号、邮箱、QQ至少填写一个"
 * )
 * public class ContactDTO {
 *     private String mobile;
 *     private String email;
 *     private String qq;
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(AtLeastOne.List.class)
@Documented
public @interface AtLeastOne {

    /**
     * 参与校验的字段名列表
     *
     * @return 字段名数组
     */
    String[] fields();

    /**
     * 至少需要填写的数量
     *
     * @return 最少数量，默认为 1
     */
    int min() default 1;

    /**
     * 错误提示消息
     *
     * @return 错误消息
     */
    String message() default "至少需要填写一个字段";

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
        AtLeastOne[] value();
    }
}
