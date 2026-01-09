package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.context.AnalysisContext;
import cn.idev.excel.metadata.data.CellData;
import cn.idev.excel.metadata.data.ImageData;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 简化版图片读取监听器
 * <p>
 * 支持读取Excel中的图片数据
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class SimpleImageReadListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();
    private final List<ErrorMessage> errorMessageList = new ArrayList<>();
    private Long lineNum = 1L;

    /**
     * 存储每个单元格的图片数据
     */
    private final Map<String, Object> imageDataCache = new HashMap<>();

    @Override
    public void invoke(Object data, AnalysisContext context) {
        lineNum++;

        try {
            // 处理图片字段
            processImageFields(data, context);

            // 进行数据验证
            Set<ConstraintViolation<Object>> violations = Validators.validate(data);
            if (!violations.isEmpty()) {
                Set<FieldError> fieldErrorSet = new HashSet<>();
                violations.forEach(violation -> {
                    try {
                        String propertyName = violation.getPropertyPath().toString();
                        Field field = data.getClass().getDeclaredField(propertyName);
                        field.setAccessible(true);

                        String fieldName = Optional.ofNullable(field.getAnnotation(ExcelProperty.class))
                                .map(anno -> anno.value()[0])
                                .orElse(field.getName());

                        String message = violation.getMessage();
                        String errorType = extractErrorType(violation);
                        Object fieldValue = null;

                        try {
                            fieldValue = field.get(data);
                        } catch (Exception ignored) {
                            // 忽略获取字段值失败
                        }

                        String fullMessage;
                        if (message != null && (message.contains(fieldName) || message.startsWith(fieldName))) {
                            fullMessage = message;
                        } else {
                            fullMessage = "【" + fieldName + "】" + message;
                        }

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
                // 设置行号
                Field[] fields = data.getClass().getDeclaredFields();
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelLine.class) && field.getType() == Long.class) {
                        try {
                            field.setAccessible(true);
                            field.set(data, lineNum);
                        } catch (IllegalAccessException e) {
                            log.error("设置行号失败: {}", e.getMessage(), e);
                        }
                    }
                }
                list.add(data);
            }
        } catch (Exception e) {
            log.error("处理Excel行数据失败: {}", e.getMessage(), e);
        }
    }

    /**
     * 处理图片字段
     * 注意：这个方法需要根据实际的Excel文件格式来调整
     * 目前的实现假设图片已经以Base64格式存储在单元格中
     */
    private void processImageFields(Object data, AnalysisContext context) {
        try {
            Class<?> clazz = data.getClass();
            Field[] fields = clazz.getDeclaredFields();

            for (Field field : fields) {
                // 检查是否有@ExcelImage注解
                if (!field.isAnnotationPresent(ExcelImage.class)) {
                    continue;
                }

                ExcelImage imageAnnotation = field.getAnnotation(ExcelImage.class);
                field.setAccessible(true);

                // 获取字段当前值
                Object currentValue = field.get(data);

                // 处理不同类型的图片字段
                Class<?> fieldType = field.getType();

                if (fieldType == String.class) {
                    // String类型：已经是Base64格式，不需要额外处理
                    if (currentValue != null && !currentValue.toString().isEmpty()) {
                        log.debug("图片字段 {} 已有数据", field.getName());
                    }
                } else if (fieldType == byte[].class) {
                    // byte[]类型：需要从Base64转换
                    if (currentValue instanceof String) {
                        String base64Data = currentValue.toString();
                        if (base64Data.startsWith("data:")) {
                            // 移除data:image/png;base64,前缀
                            String base64Only = base64Data.substring(base64Data.indexOf(",") + 1);
                            byte[] imageBytes = Base64.getDecoder().decode(base64Only);
                            field.set(data, imageBytes);
                            log.debug("转换图片字段 {} 从Base64到byte[]", field.getName());
                        }
                    }
                } else if (fieldType == List.class) {
                    // List类型：处理多张图片
                    // 初始化空列表，确保字段不为null
                    if (currentValue == null) {
                        List<String> imageList = new ArrayList<>();

                        // 如果有Excel属性注解，尝试获取对应列的值
                        ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                        if (excelProperty != null) {
                            // 这里可以通过反射获取其他字段的值来判断是否有图片数据
                            // 暂时先初始化为空列表
                            // 实际的图片数据应该通过更复杂的处理来获取
                        }

                        field.set(data, imageList);
                    }
                }
            }
        } catch (Exception e) {
            log.error("处理图片字段失败: {}", e.getMessage(), e);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed, total {} records processed", list.size());
    }

    @Override
    public List<Object> getList() {
        return list;
    }

    @Override
    public List<ErrorMessage> getErrors() {
        return errorMessageList;
    }

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