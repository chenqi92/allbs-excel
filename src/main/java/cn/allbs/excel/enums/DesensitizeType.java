package cn.allbs.excel.enums;

/**
 * 数据脱敏类型枚举
 * <p>
 * 定义了常见的数据脱敏类型及其默认脱敏规则
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
public enum DesensitizeType {

    /**
     * 手机号脱敏
     * <p>
     * 规则：保留前3位和后4位，中间4位用 * 替换
     * </p>
     * <p>示例：138****1234</p>
     */
    MOBILE_PHONE(3, 4),

    /**
     * 身份证号脱敏
     * <p>
     * 规则：保留前6位和后4位，中间用 * 替换
     * </p>
     * <p>示例：110101********1234</p>
     */
    ID_CARD(6, 4),

    /**
     * 邮箱脱敏
     * <p>
     * 规则：保留邮箱前缀第1位和@后的域名，中间用 * 替换
     * </p>
     * <p>示例：a***@example.com</p>
     */
    EMAIL(1, 0),

    /**
     * 银行卡号脱敏
     * <p>
     * 规则：保留前6位和后4位，中间用 * 替换
     * </p>
     * <p>示例：622202******1234</p>
     */
    BANK_CARD(6, 4),

    /**
     * 姓名脱敏
     * <p>
     * 规则：保留姓氏，名字用 * 替换
     * </p>
     * <p>示例：张*、欧阳**</p>
     */
    NAME(1, 0),

    /**
     * 地址脱敏
     * <p>
     * 规则：保留省市，详细地址用 * 替换
     * </p>
     * <p>示例：北京市海淀区****</p>
     */
    ADDRESS(6, 0),

    /**
     * 固定电话脱敏
     * <p>
     * 规则：保留区号和后2位，中间用 * 替换
     * </p>
     * <p>示例：010****12</p>
     */
    FIXED_PHONE(3, 2),

    /**
     * 车牌号脱敏
     * <p>
     * 规则：保留前2位和后1位，中间用 * 替换
     * </p>
     * <p>示例：京A****1</p>
     */
    CAR_LICENSE(2, 1),

    /**
     * 自定义脱敏
     * <p>
     * 使用注解中指定的 prefixKeep 和 suffixKeep 参数
     * </p>
     */
    CUSTOM(0, 0);

    /**
     * 保留前几位
     */
    private final int prefixKeep;

    /**
     * 保留后几位
     */
    private final int suffixKeep;

    DesensitizeType(int prefixKeep, int suffixKeep) {
        this.prefixKeep = prefixKeep;
        this.suffixKeep = suffixKeep;
    }

    public int getPrefixKeep() {
        return prefixKeep;
    }

    public int getSuffixKeep() {
        return suffixKeep;
    }
}

