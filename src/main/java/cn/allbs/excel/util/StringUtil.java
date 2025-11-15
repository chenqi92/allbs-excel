package cn.allbs.excel.util;

/**
 * 字符串工具类
 * 从 allbs-common 提取
 *
 * @author ChenQi
 */
public class StringUtil {

    private StringUtil() {
    }

    /**
     * 补充字符串以满足最小长度
     * <pre>
     * StringUtil.padPre("1", 3, "0")  = "001"
     * StringUtil.padPre("123", 2, "0") = "123"
     * </pre>
     *
     * @param str       字符串
     * @param minLength 最小长度
     * @param padStr    补充的字符串
     * @return 补充后的字符串
     */
    public static String padPre(String str, int minLength, String padStr) {
        if (str == null) {
            return null;
        }
        final int strLen = str.length();
        if (strLen == minLength) {
            return str;
        } else if (strLen > minLength) {
            return str;
        }

        return repeat(padStr, minLength - strLen).concat(str);
    }

    /**
     * 重复某个字符串
     *
     * @param str   字符串
     * @param count 重复次数
     * @return 重复后的字符串
     */
    public static String repeat(String str, int count) {
        if (str == null) {
            return null;
        }
        if (count <= 0) {
            return "";
        }
        if (count == 1) {
            return str;
        }

        final StringBuilder sb = new StringBuilder(str.length() * count);
        for (int i = 0; i < count; i++) {
            sb.append(str);
        }
        return sb.toString();
    }
}

