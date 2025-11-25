package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;

/**
 * 混合图片读取监听器
 * <p>
 * 结合EasyExcel和POI的优势：
 * - 使用EasyExcel读取普通数据（高性能）
 * - 使用POI读取Drawing对象图片（完整功能）
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class HybridImageReadListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();
    private final List<ErrorMessage> errorMessageList = new ArrayList<>();
    private Long lineNum = 1L;

    /**
     * POI图片提取器
     */
    private POIImageExtractor imageExtractor;

    /**
     * 是否已经初始化图片提取器
     */
    private boolean imageExtractorInitialized = false;

    /**
     * 原始输入流的字节数组（用于POI读取）
     */
    private byte[] excelBytes;

    /**
     * 当前sheet索引
     */
    private int currentSheetIndex = 0;

    /**
     * 字段到列索引的映射
     */
    private final Map<String, Integer> fieldColumnIndexMap = new HashMap<>();

    /**
     * 是否已经建立字段映射
     */
    private boolean fieldMappingBuilt = false;

    /**
     * 构造函数
     *
     * @param excelBytes Excel文件的字节数组
     */
    public HybridImageReadListener(byte[] excelBytes) {
        this.excelBytes = excelBytes;
    }

    /**
     * 默认构造函数（不支持图片读取）
     */
    public HybridImageReadListener() {
        this.excelBytes = null;
    }

    @Override
    public void invoke(Object data, AnalysisContext context) {
        lineNum++;

        try {
            // 初始化图片提取器（只执行一次）
            if (!imageExtractorInitialized && excelBytes != null) {
                initializeImageExtractor();
                imageExtractorInitialized = true;
            }

            // 建立字段映射（只执行一次）
            if (!fieldMappingBuilt) {
                buildFieldColumnIndexMap(data.getClass());
                fieldMappingBuilt = true;
            }

            // 获取当前sheet索引
            currentSheetIndex = context.readSheetHolder().getSheetNo() != null
                ? context.readSheetHolder().getSheetNo() : 0;

            // 处理图片字段
            if (imageExtractor != null) {
                processImageFields(data, (int)(lineNum - 2)); // 减2是因为lineNum从1开始，且有表头
            }

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
     * 初始化图片提取器
     */
    private void initializeImageExtractor() {
        try {
            imageExtractor = new POIImageExtractor();
            InputStream is = new ByteArrayInputStream(excelBytes);
            boolean success = imageExtractor.extractImages(is, "excel.xlsx");

            if (success) {
                log.info("成功初始化POI图片提取器，提取了 {} 个sheet的图片",
                        imageExtractor.getExtractedImages().size());
            } else {
                log.warn("POI图片提取器初始化失败");
                imageExtractor = null;
            }
        } catch (Exception e) {
            log.error("初始化图片提取器失败", e);
            imageExtractor = null;
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
     */
    private void processImageFields(Object data, int rowIndex) {
        try {
            Class<?> clazz = data.getClass();
            Field[] fields = clazz.getDeclaredFields();

            for (Field field : fields) {
                if (!field.isAnnotationPresent(ExcelImage.class)) {
                    continue;
                }

                Integer columnIndex = fieldColumnIndexMap.get(field.getName());
                if (columnIndex == null) {
                    continue;
                }

                // 从POI提取器获取图片
                List<POIImageExtractor.ImageData> images =
                    imageExtractor.getImages(currentSheetIndex, rowIndex, columnIndex);

                if (images.isEmpty()) {
                    continue;
                }

                field.setAccessible(true);
                Class<?> fieldType = field.getType();

                if (fieldType == String.class) {
                    // String类型：存储第一张图片的Base64
                    POIImageExtractor.ImageData imageData = images.get(0);
                    field.set(data, imageData.toBase64());
                    log.debug("设置图片字段 {} (String)，大小: {} bytes",
                            field.getName(), imageData.getData().length);

                } else if (fieldType == byte[].class) {
                    // byte[]类型：存储原始字节数据
                    POIImageExtractor.ImageData imageData = images.get(0);
                    field.set(data, imageData.getData());
                    log.debug("设置图片字段 {} (byte[])，大小: {} bytes",
                            field.getName(), imageData.getData().length);

                } else if (fieldType == List.class) {
                    // List类型：存储多张图片
                    Type genericType = field.getGenericType();
                    if (genericType instanceof ParameterizedType) {
                        ParameterizedType pt = (ParameterizedType) genericType;
                        Type[] actualTypes = pt.getActualTypeArguments();

                        if (actualTypes.length > 0 && actualTypes[0] == String.class) {
                            // List<String>类型
                            List<String> imageList = new ArrayList<>();
                            for (POIImageExtractor.ImageData imageData : images) {
                                imageList.add(imageData.toBase64());
                            }

                            // 检查相邻列是否有更多图片
                            for (int i = 1; i < 5; i++) {
                                List<POIImageExtractor.ImageData> extraImages =
                                    imageExtractor.getImages(currentSheetIndex, rowIndex, columnIndex + i);
                                if (!extraImages.isEmpty()) {
                                    for (POIImageExtractor.ImageData imageData : extraImages) {
                                        imageList.add(imageData.toBase64());
                                    }
                                }
                            }

                            field.set(data, imageList);
                            log.debug("设置图片字段 {} (List<String>)，共 {} 张图片",
                                    field.getName(), imageList.size());
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("处理图片字段失败", e);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed, total {} records with images processed", list.size());

        // 清理资源
        if (imageExtractor != null) {
            imageExtractor.clear();
            imageExtractor = null;
        }
        excelBytes = null;
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