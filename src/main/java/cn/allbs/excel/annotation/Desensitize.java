package cn.allbs.excel.annotation;

import cn.allbs.excel.enums.DesensitizeType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 数据脱敏注解
 * <p>
 * 用于标记需要进行数据脱敏的字段，仅在导出时生效
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty("手机号")
 * &#64;Desensitize(type = DesensitizeType.MOBILE_PHONE)
 * private String phone;
 *
 * &#64;ExcelProperty("身份证")
 * &#64;Desensitize(type = DesensitizeType.ID_CARD)
 * private String idCard;
 *
 * &#64;ExcelProperty("自定义")
 * &#64;Desensitize(type = DesensitizeType.CUSTOM, prefixKeep = 2, suffixKeep = 3)
 * private String custom;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Desensitize {

    /**
     * 脱敏类型
     * <p>
     * 支持的类型：
     * <ul>
     *     <li>MOBILE_PHONE - 手机号：138****1234</li>
     *     <li>ID_CARD - 身份证：110101********1234</li>
     *     <li>EMAIL - 邮箱：a***@example.com</li>
     *     <li>BANK_CARD - 银行卡：622202******1234</li>
     *     <li>NAME - 姓名：张*</li>
     *     <li>ADDRESS - 地址：北京市海淀区****</li>
     *     <li>FIXED_PHONE - 固定电话：010****12</li>
     *     <li>CAR_LICENSE - 车牌号：京A****1</li>
     *     <li>CUSTOM - 自定义：使用 prefixKeep 和 suffixKeep 参数</li>
     * </ul>
     * </p>
     *
     * @return 脱敏类型
     */
    DesensitizeType type();

    /**
     * 保留前几位
     * <p>
     * 仅在 type = CUSTOM 时生效
     * </p>
     *
     * @return 保留前几位，默认 3
     */
    int prefixKeep() default 3;

    /**
     * 保留后几位
     * <p>
     * 仅在 type = CUSTOM 时生效
     * </p>
     *
     * @return 保留后几位，默认 4
     */
    int suffixKeep() default 4;

    /**
     * 脱敏字符
     * <p>
     * 用于替换敏感信息的字符
     * </p>
     *
     * @return 脱敏字符，默认 *
     */
    String maskChar() default "*";

    /**
     * 是否启用脱敏
     * <p>
     * 可以通过此参数动态控制是否启用脱敏，方便在不同环境下切换
     * </p>
     *
     * @return 是否启用脱敏，默认 true
     */
    boolean enabled() default true;
}

