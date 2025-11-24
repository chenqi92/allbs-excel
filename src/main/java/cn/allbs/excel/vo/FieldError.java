package cn.allbs.excel.vo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 字段验证错误详情
 *
 * @author ChenQi
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class FieldError {

    /**
     * 字段名称（Excel 列名）
     */
    private String fieldName;

    /**
     * 字段属性名（Java 属性名）
     */
    private String propertyName;

    /**
     * 错误类型
     * <p>
     * 常见类型：
     * - NotNull: 不能为空
     * - NotBlank: 不能为空白
     * - Size: 长度不符合要求
     * - Min/Max: 数值范围不符合
     * - Email: 邮箱格式不正确
     * - Pattern: 正则表达式不匹配
     * - DecimalMin/DecimalMax: 小数范围不符合
     * - Digits: 数字格式不正确
     * - Past/Future: 日期不符合要求
     * - AssertTrue/AssertFalse: 布尔值不符合要求
     * </p>
     */
    private String errorType;

    /**
     * 错误消息（原始消息，不含字段名）
     */
    private String message;

    /**
     * 完整错误消息（包含字段名的格式化消息）
     */
    private String fullMessage;

    /**
     * 字段值（可选，用于调试）
     */
    private Object fieldValue;

    /**
     * 约束注解的详细信息（可选）
     */
    private String constraintDetails;

    /**
     * 获取默认格式的错误消息
     * 格式：【字段名】错误消息
     *
     * @return 格式化的错误消息
     */
    public String getDefaultMessage() {
        if (fullMessage != null && !fullMessage.isEmpty()) {
            return fullMessage;
        }
        return "【" + fieldName + "】" + message;
    }

    /**
     * 获取简洁的错误消息
     * 格式：字段名: 错误消息
     *
     * @return 简洁的错误消息
     */
    public String getSimpleMessage() {
        return fieldName + ": " + message;
    }

    /**
     * 获取详细的错误消息（包含错误类型）
     * 格式：【字段名】[错误类型] 错误消息
     *
     * @return 详细的错误消息
     */
    public String getDetailedMessage() {
        return "【" + fieldName + "】[" + errorType + "] " + message;
    }

    /**
     * 判断是否为必填验证错误
     *
     * @return true 如果是 NotNull、NotBlank、NotEmpty 类型的错误
     */
    public boolean isRequiredError() {
        return "NotNull".equals(errorType)
            || "NotBlank".equals(errorType)
            || "NotEmpty".equals(errorType);
    }

    /**
     * 判断是否为格式验证错误
     *
     * @return true 如果是 Email、Pattern、Digits 等格式类型的错误
     */
    public boolean isFormatError() {
        return "Email".equals(errorType)
            || "Pattern".equals(errorType)
            || "Digits".equals(errorType);
    }

    /**
     * 判断是否为范围验证错误
     *
     * @return true 如果是 Size、Min、Max、DecimalMin、DecimalMax 类型的错误
     */
    public boolean isRangeError() {
        return "Size".equals(errorType)
            || "Min".equals(errorType)
            || "Max".equals(errorType)
            || "DecimalMin".equals(errorType)
            || "DecimalMax".equals(errorType);
    }
}
