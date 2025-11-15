package cn.allbs.excel.util;

import cn.allbs.excel.enums.DesensitizeType;

/**
 * 数据脱敏工具类
 * <p>
 * 提供各种常见数据类型的脱敏方法
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
public class DesensitizeUtil {

    private DesensitizeUtil() {
    }

    /**
     * 通用脱敏方法
     *
     * @param value      原始值
     * @param type       脱敏类型
     * @param prefixKeep 保留前几位（仅 CUSTOM 类型使用）
     * @param suffixKeep 保留后几位（仅 CUSTOM 类型使用）
     * @param maskChar   脱敏字符
     * @return 脱敏后的值
     */
    public static String desensitize(String value, DesensitizeType type, int prefixKeep, int suffixKeep, String maskChar) {
        if (value == null || value.isEmpty()) {
            return value;
        }

        switch (type) {
            case MOBILE_PHONE:
                return desensitizeMobilePhone(value, maskChar);
            case ID_CARD:
                return desensitizeIdCard(value, maskChar);
            case EMAIL:
                return desensitizeEmail(value, maskChar);
            case BANK_CARD:
                return desensitizeBankCard(value, maskChar);
            case NAME:
                return desensitizeName(value, maskChar);
            case ADDRESS:
                return desensitizeAddress(value, maskChar);
            case FIXED_PHONE:
                return desensitizeFixedPhone(value, maskChar);
            case CAR_LICENSE:
                return desensitizeCarLicense(value, maskChar);
            case CUSTOM:
                return desensitizeCustom(value, prefixKeep, suffixKeep, maskChar);
            default:
                return value;
        }
    }

    /**
     * 手机号脱敏
     * <p>
     * 规则：保留前3位和后4位，中间4位用 * 替换
     * </p>
     * <p>示例：13812345678 -> 138****5678</p>
     */
    public static String desensitizeMobilePhone(String phone, String maskChar) {
        if (phone == null || phone.length() != 11) {
            return phone;
        }
        return phone.substring(0, 3) + repeat(maskChar, 4) + phone.substring(7);
    }

    /**
     * 身份证号脱敏
     * <p>
     * 规则：保留前6位和后4位，中间用 * 替换
     * </p>
     * <p>示例：110101199001011234 -> 110101********1234</p>
     */
    public static String desensitizeIdCard(String idCard, String maskChar) {
        if (idCard == null || (idCard.length() != 15 && idCard.length() != 18)) {
            return idCard;
        }
        int maskLength = idCard.length() - 10;
        return idCard.substring(0, 6) + repeat(maskChar, maskLength) + idCard.substring(idCard.length() - 4);
    }

    /**
     * 邮箱脱敏
     * <p>
     * 规则：保留邮箱前缀第1位和@后的域名，中间用 * 替换
     * </p>
     * <p>示例：example@gmail.com -> e***@gmail.com</p>
     */
    public static String desensitizeEmail(String email, String maskChar) {
        if (email == null || !email.contains("@")) {
            return email;
        }
        int atIndex = email.indexOf("@");
        if (atIndex <= 1) {
            return email;
        }
        String prefix = email.substring(0, 1);
        String domain = email.substring(atIndex);
        return prefix + repeat(maskChar, 3) + domain;
    }

    /**
     * 银行卡号脱敏
     * <p>
     * 规则：保留前6位和后4位，中间用 * 替换
     * </p>
     * <p>示例：6222021234567890123 -> 622202*********0123</p>
     */
    public static String desensitizeBankCard(String bankCard, String maskChar) {
        if (bankCard == null || bankCard.length() < 10) {
            return bankCard;
        }
        int maskLength = bankCard.length() - 10;
        return bankCard.substring(0, 6) + repeat(maskChar, maskLength) + bankCard.substring(bankCard.length() - 4);
    }

    /**
     * 姓名脱敏
     * <p>
     * 规则：保留姓氏，名字用 * 替换
     * </p>
     * <p>示例：张三 -> 张*，欧阳锋 -> 欧阳*</p>
     */
    public static String desensitizeName(String name, String maskChar) {
        if (name == null || name.isEmpty()) {
            return name;
        }
        int length = name.length();
        if (length == 1) {
            return name;
        }
        // 复姓处理：欧阳、司马、上官等
        if (length >= 3 && isCompoundSurname(name.substring(0, 2))) {
            return name.substring(0, 2) + repeat(maskChar, length - 2);
        }
        // 普通姓名
        return name.substring(0, 1) + repeat(maskChar, length - 1);
    }

    /**
     * 地址脱敏
     * <p>
     * 规则：保留前6位（通常是省市），详细地址用 * 替换
     * </p>
     * <p>示例：北京市海淀区中关村大街1号 -> 北京市海淀区****</p>
     */
    public static String desensitizeAddress(String address, String maskChar) {
        if (address == null || address.length() <= 6) {
            return address;
        }
        return address.substring(0, 6) + repeat(maskChar, 4);
    }

    /**
     * 固定电话脱敏
     * <p>
     * 规则：保留区号和后2位，中间用 * 替换
     * </p>
     * <p>示例：010-12345678 -> 010****78</p>
     */
    public static String desensitizeFixedPhone(String phone, String maskChar) {
        if (phone == null || phone.isEmpty()) {
            return phone;
        }
        // 移除分隔符
        String cleanPhone = phone.replaceAll("[-\\s]", "");
        if (cleanPhone.length() < 5) {
            return phone;
        }
        // 保留前3位（区号）和后2位
        int maskLength = cleanPhone.length() - 5;
        return cleanPhone.substring(0, 3) + repeat(maskChar, maskLength) + cleanPhone.substring(cleanPhone.length() - 2);
    }

    /**
     * 车牌号脱敏
     * <p>
     * 规则：保留前2位和后1位，中间用 * 替换
     * </p>
     * <p>示例：京A12345 -> 京A****5</p>
     */
    public static String desensitizeCarLicense(String carLicense, String maskChar) {
        if (carLicense == null || carLicense.length() < 3) {
            return carLicense;
        }
        int maskLength = carLicense.length() - 3;
        return carLicense.substring(0, 2) + repeat(maskChar, maskLength) + carLicense.substring(carLicense.length() - 1);
    }

    /**
     * 自定义脱敏
     * <p>
     * 根据指定的前后保留位数进行脱敏
     * </p>
     *
     * @param value      原始值
     * @param prefixKeep 保留前几位
     * @param suffixKeep 保留后几位
     * @param maskChar   脱敏字符
     * @return 脱敏后的值
     */
    public static String desensitizeCustom(String value, int prefixKeep, int suffixKeep, String maskChar) {
        if (value == null || value.isEmpty()) {
            return value;
        }
        int length = value.length();
        if (length <= prefixKeep + suffixKeep) {
            return value;
        }
        int maskLength = length - prefixKeep - suffixKeep;
        String prefix = prefixKeep > 0 ? value.substring(0, prefixKeep) : "";
        String suffix = suffixKeep > 0 ? value.substring(length - suffixKeep) : "";
        return prefix + repeat(maskChar, maskLength) + suffix;
    }

    /**
     * 重复字符串
     *
     * @param str   要重复的字符串
     * @param count 重复次数
     * @return 重复后的字符串
     */
    private static String repeat(String str, int count) {
        if (count <= 0) {
            return "";
        }
        StringBuilder sb = new StringBuilder(str.length() * count);
        for (int i = 0; i < count; i++) {
            sb.append(str);
        }
        return sb.toString();
    }

    /**
     * 判断是否为复姓
     */
    private static boolean isCompoundSurname(String surname) {
        String[] compoundSurnames = {
                "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫", "万俟", "闻人",
                "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台", "皇甫", "宗政", "濮阳", "公冶",
                "太叔", "申屠", "公孙", "慕容", "仲孙", "钟离", "长孙", "宇文", "司徒", "鲜于",
                "司空", "闾丘", "子车", "亓官", "司寇", "巫马", "公西", "颛孙", "壤驷", "公良",
                "漆雕", "乐正", "宰父", "谷梁", "拓跋", "夹谷", "轩辕", "令狐", "段干", "百里",
                "呼延", "东郭", "南门", "羊舌", "微生", "公户", "公玉", "公仪", "梁丘", "公仲",
                "公上", "公门", "公山", "公坚", "左丘", "公伯", "西门", "公祖", "第五", "公乘",
                "贯丘", "公皙", "南荣", "东里", "东宫", "仲长", "子书", "子桑", "即墨", "达奚",
                "褚师"
        };
        for (String compoundSurname : compoundSurnames) {
            if (compoundSurname.equals(surname)) {
                return true;
            }
        }
        return false;
    }
}

