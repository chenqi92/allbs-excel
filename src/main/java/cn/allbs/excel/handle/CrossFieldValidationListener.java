package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.validation.CrossFieldValidators;
import cn.allbs.excel.validation.CrossFieldValidators.UniqueValidator;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.context.AnalysisContext;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 支持跨字段校验的 AnalysisEventListener
 * <p>
 * 在 {@link DefaultAnalysisEventListener} 基础上，增加了跨字段校验功能
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
@Slf4j
public class CrossFieldValidationListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();
    private final List<ErrorMessage> errorMessageList = new ArrayList<>();
    private Long lineNum = 1L;

    /**
     * 校验分组
     */
    private final Class<?>[] validationGroups;

    /**
     * 是否启用跨字段校验
     */
    private final boolean enableCrossValidation;

    /**
     * 唯一性校验器（懒加载）
     */
    private UniqueValidator uniqueValidator;
    private Class<?> dataClass;

    public CrossFieldValidationListener() {
        this(true);
    }

    public CrossFieldValidationListener(boolean enableCrossValidation) {
        this(enableCrossValidation, (Class<?>[]) null);
    }

    public CrossFieldValidationListener(boolean enableCrossValidation, Class<?>... validationGroups) {
        this.enableCrossValidation = enableCrossValidation;
        this.validationGroups = validationGroups;
    }

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        lineNum++;

        // 初始化唯一性校验器
        if (enableCrossValidation && dataClass == null) {
            dataClass = o.getClass();
            uniqueValidator = CrossFieldValidators.createUniqueValidator(dataClass);
        }

        Set<FieldError> fieldErrorSet = new HashSet<>();

        // 1. 单字段 JSR-380 校验
        Set<ConstraintViolation<Object>> violations = Validators.validate(o);
        if (!violations.isEmpty()) {
            violations.forEach(violation -> {
                try {
                    fieldErrorSet.add(buildFieldError(o, violation));
                } catch (Exception e) {
                    log.error("字段解析错误: {}", e.getMessage(), e);
                }
            });
        }

        // 2. 跨字段校验（新增）
        if (enableCrossValidation) {
            // 行内校验
            List<FieldError> crossErrors = CrossFieldValidators.validate(o, validationGroups);
            fieldErrorSet.addAll(crossErrors);

            // 跨行唯一性校验
            if (uniqueValidator != null && uniqueValidator.hasRules()) {
                List<FieldError> uniqueErrors = uniqueValidator.validate(o, lineNum.intValue(), validationGroups);
                fieldErrorSet.addAll(uniqueErrors);
            }
        }

        if (!fieldErrorSet.isEmpty()) {
            errorMessageList.add(new ErrorMessage(lineNum, fieldErrorSet));
        } else {
            // 设置行号
            setLineNumber(o);
            list.add(o);
        }
    }

    /**
     * 构建字段错误信息
     */
    private FieldError buildFieldError(Object o, ConstraintViolation<Object> violation) {
        String propertyName = violation.getPropertyPath().toString();
        // 使用 FieldAccessorCache 获取字段，支持父类字段遍历
        Field field = FieldAccessorCache.getField(o.getClass(), propertyName);

        if (field == null) {
            // 无法获取字段信息时使用属性名作为字段名
            return FieldError.builder()
                    .fieldName(propertyName)
                    .propertyName(propertyName)
                    .errorType(extractErrorType(violation))
                    .message(violation.getMessage())
                    .fullMessage("【" + propertyName + "】" + violation.getMessage())
                    .build();
        }

        field.setAccessible(true);

        // 获取字段名称（Excel 列名）
        String fieldName = Optional.ofNullable(field.getAnnotation(ExcelProperty.class))
                .map(anno -> anno.value()[0])
                .orElse(field.getName());

        // 获取错误消息
        String message = violation.getMessage();

        // 获取错误类型
        String errorType = extractErrorType(violation);

        // 获取字段值
        Object fieldValue = null;
        try {
            fieldValue = field.get(o);
        } catch (Exception ignored) {
        }

        // 智能拼接完整错误消息
        String fullMessage;
        if (message != null && (message.contains(fieldName) || message.startsWith(fieldName))) {
            fullMessage = message;
        } else {
            fullMessage = "【" + fieldName + "】" + message;
        }

        return FieldError.builder()
                .fieldName(fieldName)
                .propertyName(propertyName)
                .errorType(errorType)
                .message(message)
                .fullMessage(fullMessage)
                .fieldValue(fieldValue)
                .build();
    }

    /**
     * 设置行号
     */
    private void setLineNumber(Object o) {
        Field[] fields = o.getClass().getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelLine.class) && field.getType() == Long.class) {
                try {
                    field.setAccessible(true);
                    field.set(o, lineNum);
                } catch (IllegalAccessException e) {
                    log.error("设置行号失败: {}", e.getMessage(), e);
                }
            }
        }
    }

    /**
     * 从约束违反对象中提取错误类型
     */
    private String extractErrorType(ConstraintViolation<?> violation) {
        try {
            ConstraintDescriptor<?> descriptor = violation.getConstraintDescriptor();
            Class<?> annotationType = descriptor.getAnnotation().annotationType();
            return annotationType.getSimpleName();
        } catch (Exception e) {
            log.debug("无法提取错误类型: {}", e.getMessage());
            return "Unknown";
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed with cross-field validation");
    }

    @Override
    public List<Object> getList() {
        return list;
    }

    @Override
    public List<ErrorMessage> getErrors() {
        return errorMessageList;
    }

    /**
     * 重置校验器状态（用于多次导入同一文件或测试）
     */
    public void reset() {
        list.clear();
        errorMessageList.clear();
        lineNum = 1L;
        if (uniqueValidator != null) {
            uniqueValidator.reset();
        }
    }
}
