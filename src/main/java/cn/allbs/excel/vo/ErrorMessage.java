package cn.allbs.excel.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.HashSet;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 校验错误信息
 *
 * @author ChenQi
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ErrorMessage {

    /**
     * 行号
     */
    private Long lineNum;

    /**
     * 错误信息（简单字符串格式，保留用于向后兼容）
     * @deprecated 推荐使用 fieldErrors 获取结构化的错误信息
     */
    @Deprecated
    private Set<String> errors = new HashSet<>();

    /**
     * 字段错误详情（结构化的错误信息）
     */
    private Set<FieldError> fieldErrors = new HashSet<>();

    public ErrorMessage(Set<String> errors) {
        this.errors = errors;
    }

    public ErrorMessage(String error) {
        HashSet<String> objects = new HashSet<>();
        objects.add(error);
        this.errors = objects;
    }

    /**
     * 构造函数 - 使用结构化错误信息
     *
     * @param lineNum 行号
     * @param fieldErrors 字段错误集合
     */
    public ErrorMessage(Long lineNum, Set<FieldError> fieldErrors) {
        this.lineNum = lineNum;
        this.fieldErrors = fieldErrors;
        // 同时填充 errors 字段以保持向后兼容
        this.errors = fieldErrors.stream()
                .map(FieldError::getDefaultMessage)
                .collect(Collectors.toSet());
    }

    /**
     * 添加字段错误
     *
     * @param fieldError 字段错误
     */
    public void addFieldError(FieldError fieldError) {
        this.fieldErrors.add(fieldError);
        this.errors.add(fieldError.getDefaultMessage());
    }

    /**
     * 获取所有错误消息（默认格式）
     *
     * @return 错误消息集合
     */
    public Set<String> getErrorMessages() {
        if (fieldErrors != null && !fieldErrors.isEmpty()) {
            return fieldErrors.stream()
                    .map(FieldError::getDefaultMessage)
                    .collect(Collectors.toSet());
        }
        return errors;
    }

    /**
     * 获取所有错误消息（简洁格式）
     *
     * @return 简洁的错误消息集合
     */
    public Set<String> getSimpleErrorMessages() {
        if (fieldErrors != null && !fieldErrors.isEmpty()) {
            return fieldErrors.stream()
                    .map(FieldError::getSimpleMessage)
                    .collect(Collectors.toSet());
        }
        return errors;
    }

    /**
     * 获取所有错误消息（详细格式，包含错误类型）
     *
     * @return 详细的错误消息集合
     */
    public Set<String> getDetailedErrorMessages() {
        if (fieldErrors != null && !fieldErrors.isEmpty()) {
            return fieldErrors.stream()
                    .map(FieldError::getDetailedMessage)
                    .collect(Collectors.toSet());
        }
        return errors;
    }

    /**
     * 获取指定字段的错误
     *
     * @param fieldName 字段名称
     * @return 字段错误集合
     */
    public Set<FieldError> getFieldErrorsByName(String fieldName) {
        return fieldErrors.stream()
                .filter(fe -> fieldName.equals(fe.getFieldName()))
                .collect(Collectors.toSet());
    }

    /**
     * 获取必填错误
     *
     * @return 必填错误集合
     */
    public Set<FieldError> getRequiredErrors() {
        return fieldErrors.stream()
                .filter(FieldError::isRequiredError)
                .collect(Collectors.toSet());
    }

    /**
     * 获取格式错误
     *
     * @return 格式错误集合
     */
    public Set<FieldError> getFormatErrors() {
        return fieldErrors.stream()
                .filter(FieldError::isFormatError)
                .collect(Collectors.toSet());
    }

    /**
     * 获取范围错误
     *
     * @return 范围错误集合
     */
    public Set<FieldError> getRangeErrors() {
        return fieldErrors.stream()
                .filter(FieldError::isRangeError)
                .collect(Collectors.toSet());
    }

    /**
     * 判断是否有错误
     *
     * @return true 如果有任何错误
     */
    public boolean hasErrors() {
        return (errors != null && !errors.isEmpty())
            || (fieldErrors != null && !fieldErrors.isEmpty());
    }

    /**
     * 获取错误数量
     *
     * @return 错误数量
     */
    public int getErrorCount() {
        return fieldErrors != null ? fieldErrors.size() : errors.size();
    }
}
