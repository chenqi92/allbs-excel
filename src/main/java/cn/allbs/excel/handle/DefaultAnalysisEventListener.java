package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 默认的 AnalysisEventListener
 *
 * @author ChenQi
 */
@Slf4j
public class DefaultAnalysisEventListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();

    private final List<ErrorMessage> errorMessageList = new ArrayList<>();

    private Long lineNum = 1L;

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        lineNum++;

        Set<ConstraintViolation<Object>> violations = Validators.validate(o);
        if (!violations.isEmpty()) {

            Set<FieldError> fieldErrorSet = new HashSet<>();
            violations.forEach(violation -> {
                try {
                    String propertyName = violation.getPropertyPath().toString();
                    Field field = o.getClass().getDeclaredField(propertyName);
                    field.setAccessible(true);

                    // 获取字段名称（Excel 列名）
                    String fieldName = Optional.ofNullable(field.getAnnotation(ExcelProperty.class))
                            .map(anno -> anno.value()[0])
                            .orElse(field.getName());

                    // 获取错误消息
                    String message = violation.getMessage();

                    // 获取错误类型（从约束注解中提取）
                    String errorType = extractErrorType(violation);

                    // 获取字段值
                    Object fieldValue = null;
                    try {
                        fieldValue = field.get(o);
                    } catch (Exception ignored) {
                        // 获取字段值失败，不影响主流程
                    }

                    // 智能拼接完整错误消息：如果消息中已包含字段名，则不重复添加
                    String fullMessage;
                    if (message != null && (message.contains(fieldName) || message.startsWith(fieldName))) {
                        fullMessage = message;
                    } else {
                        fullMessage = "【" + fieldName + "】" + message;
                    }

                    // 构建 FieldError 对象
                    FieldError fieldError = FieldError.builder()
                            .fieldName(fieldName)
                            .propertyName(propertyName)
                            .errorType(errorType)
                            .message(message)
                            .fullMessage(fullMessage)
                            .fieldValue(fieldValue)
                            .build();

                    fieldErrorSet.add(fieldError);
                } catch (Exception e) {
                    log.error("字段解析错误: {}", e.getMessage(), e);
                }
            });
            errorMessageList.add(new ErrorMessage(lineNum, fieldErrorSet));
        } else {
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
            list.add(o);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed");
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
     * 从约束违反对象中提取错误类型
     *
     * @param violation 约束违反对象
     * @return 错误类型名称
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

}
