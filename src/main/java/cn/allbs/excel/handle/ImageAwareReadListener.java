package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.ImageData;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;

/**
 * 图片感知的读取监听器
 * <p>
 * 支持读取Excel中的图片数据，包括：
 * 1. 单元格中的Base64文本
 * 2. 单元格中的ImageData（如果EasyExcel支持）
 * 注意：由于EasyExcel 4.x的限制，可能无法直接读取Drawing对象
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class ImageAwareReadListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();
    private final List<ErrorMessage> errorMessageList = new ArrayList<>();
    private Long lineNum = 1L;

    /**
     * 字段名到列索引的映射
     */
    private final Map<String, Integer> fieldColumnIndexMap = new HashMap<>();

    /**
     * 是否已经建立了字段映射
     */
    private boolean fieldMappingBuilt = false;

    @Override
    public void invoke(Object data, AnalysisContext context) {
        lineNum++;

        try {
            // 第一次调用时，建立字段到列索引的映射
            if (!fieldMappingBuilt) {
                buildFieldColumnIndexMap(data.getClass());
                fieldMappingBuilt = true;
            }

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
     * 构建字段到列索引的映射
     */
    private void buildFieldColumnIndexMap(Class<?> clazz) {
        try {
            Field[] fields = clazz.getDeclaredFields();
            int columnIndex = 0;

            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelProperty.class)) {
                    fieldColumnIndexMap.put(field.getName(), columnIndex);
                    columnIndex++;
                }
            }

            log.debug("字段列索引映射: {}", fieldColumnIndexMap);
        } catch (Exception e) {
            log.error("构建字段列索引映射失败", e);
        }
    }

    /**
     * 处理图片字段
     * 这个方法主要处理Base64格式的图片数据
     * 对于真实的Drawing对象，由于EasyExcel的限制，可能需要其他方式处理
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
                    // String类型：检查是否需要处理
                    processStringImageField(field, data, currentValue);
                } else if (fieldType == byte[].class) {
                    // byte[]类型：需要从Base64转换
                    processByteArrayImageField(field, data, currentValue);
                } else if (fieldType == List.class) {
                    // List类型：处理多张图片
                    processListImageField(field, data, currentValue);
                }
            }
        } catch (Exception e) {
            log.error("处理图片字段失败: {}", e.getMessage(), e);
        }
    }

    /**
     * 处理String类型的图片字段
     */
    private void processStringImageField(Field field, Object data, Object currentValue) throws IllegalAccessException {
        if (currentValue instanceof String) {
            String strValue = (String) currentValue;

            // 如果是纯Base64（没有data:前缀），添加默认前缀
            if (strValue != null && !strValue.isEmpty() && !strValue.startsWith("data:")) {
                // 检查是否像是Base64
                if (strValue.matches("^[A-Za-z0-9+/]+=*$")) {
                    strValue = "data:image/png;base64," + strValue;
                    field.set(data, strValue);
                    log.debug("为字段 {} 添加Base64前缀", field.getName());
                }
            }
        }
    }

    /**
     * 处理byte[]类型的图片字段
     */
    private void processByteArrayImageField(Field field, Object data, Object currentValue) throws IllegalAccessException {
        // 如果字段是null，但是有对应的String字段包含Base64数据，可以转换
        if (currentValue == null) {
            // 尝试查找同名的String字段（去掉Bytes后缀）
            String fieldName = field.getName();
            if (fieldName.endsWith("Bytes") || fieldName.endsWith("Data")) {
                String baseFieldName = fieldName.substring(0, fieldName.lastIndexOf("Bytes"));
                try {
                    Field stringField = data.getClass().getDeclaredField(baseFieldName);
                    stringField.setAccessible(true);
                    Object stringValue = stringField.get(data);

                    if (stringValue instanceof String) {
                        String base64Data = (String) stringValue;
                        if (base64Data != null && base64Data.startsWith("data:")) {
                            // 移除data:image/png;base64,前缀
                            String base64Only = base64Data.substring(base64Data.indexOf(",") + 1);
                            byte[] imageBytes = Base64.getDecoder().decode(base64Only);
                            field.set(data, imageBytes);
                            log.debug("转换图片字段 {} 从Base64到byte[]", field.getName());
                        }
                    }
                } catch (NoSuchFieldException ignored) {
                    // 没有对应的String字段，忽略
                }
            }
        }
    }

    /**
     * 处理List类型的图片字段
     */
    private void processListImageField(Field field, Object data, Object currentValue) throws IllegalAccessException {
        if (currentValue == null) {
            // 初始化为空列表
            Type genericType = field.getGenericType();
            if (genericType instanceof ParameterizedType) {
                ParameterizedType pt = (ParameterizedType) genericType;
                Type[] actualTypes = pt.getActualTypeArguments();

                if (actualTypes.length > 0 && actualTypes[0] == String.class) {
                    field.set(data, new ArrayList<String>());
                    log.debug("初始化字段 {} 为空列表", field.getName());
                }
            }
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