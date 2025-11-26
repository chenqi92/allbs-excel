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
 * - 自动兼容：优先读取Drawing图片，如果没有则使用单元格中的Base64文本
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
     * 字段到列索引的映射（字段名 -> 列索引）
     */
    private final Map<String, Integer> fieldColumnIndexMap = new HashMap<>();

    /**
     * 是否已经建立字段映射
     */
    private boolean fieldMappingBuilt = false;

    /**
     * 缓存的表头名称到列索引映射（从 EasyExcel invokeHeadMap 回调获取）
     */
    private Map<Integer, String> headNameMap = new HashMap<>();

    /**
     * 重写 invokeHeadMap 回调，在解析表头时捕获映射关系
     * <p>
     * 这是获取表头映射最可靠的方式，EasyExcel 会在读取表头后自动调用此方法
     * </p>
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        this.headNameMap = new HashMap<>(headMap);
        log.info("从 invokeHeadMap 回调获取表头映射: {}", headNameMap);
    }

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
                // 从 EasyExcel 的 context 获取表头映射
                buildFieldColumnIndexMapFromContext(data.getClass(), context);
                fieldMappingBuilt = true;
            }

            // 获取当前sheet索引
            currentSheetIndex = context.readSheetHolder().getSheetNo() != null
                ? context.readSheetHolder().getSheetNo() : 0;

            // 处理图片字段
            // POI 中 Row 索引从0开始，表头在 Row=0，第一条数据在 Row=1
            // lineNum 初始为1，第一次 invoke 后 lineNum=2，所以 rowIndex = lineNum - 1 = 1
            if (imageExtractor != null) {
                processImageFields(data, (int)(lineNum - 1));
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
     * 从 EasyExcel 的 AnalysisContext 构建字段到列索引的映射
     * <p>
     * 使用 invokeHeadMap 回调缓存的表头映射（headNameMap）
     * 这是最可靠的方式，因为 EasyExcel 在解析表头后会自动调用 invokeHeadMap
     * </p>
     */
    private void buildFieldColumnIndexMapFromContext(Class<?> clazz, AnalysisContext context) {
        try {
            // headNameMap 已在 invokeHeadMap 回调中设置
            if (headNameMap.isEmpty()) {
                log.warn("表头映射为空，可能 invokeHeadMap 尚未被调用");
            } else {
                log.debug("使用 invokeHeadMap 回调获取的表头映射: {}", headNameMap);
            }

            // 获取类的所有字段
            Field[] fields = clazz.getDeclaredFields();

            // 建立字段名到 @ExcelProperty.value 的映射
            Map<String, Field> headNameToField = new HashMap<>();
            Map<String, Integer> explicitIndexFields = new HashMap<>();

            for (Field field : fields) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                if (excelProperty != null) {
                    // 如果有显式指定 index，直接使用
                    if (excelProperty.index() >= 0) {
                        explicitIndexFields.put(field.getName(), excelProperty.index());
                        fieldColumnIndexMap.put(field.getName(), excelProperty.index());
                        log.debug("字段 {} 使用显式索引: {}", field.getName(), excelProperty.index());
                    } else {
                        // 通过表头名称匹配
                        String[] headNames = excelProperty.value();
                        if (headNames.length > 0 && !headNames[0].isEmpty()) {
                            headNameToField.put(headNames[0], field);
                        }
                    }
                }
            }

            // 根据 EasyExcel 的表头映射，为没有显式 index 的字段分配列索引
            if (headNameMap != null) {
                for (Map.Entry<Integer, String> entry : headNameMap.entrySet()) {
                    Integer colIndex = entry.getKey();
                    String headName = entry.getValue();

                    Field matchedField = headNameToField.get(headName);
                    if (matchedField != null && !fieldColumnIndexMap.containsKey(matchedField.getName())) {
                        fieldColumnIndexMap.put(matchedField.getName(), colIndex);
                        log.debug("字段 {} 通过表头 '{}' 匹配到列索引: {}", matchedField.getName(), headName, colIndex);
                    }
                }
            }

            // 对于仍未映射的字段，按声明顺序分配（作为后备方案）
            int autoIndex = 0;
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelProperty.class) && !fieldColumnIndexMap.containsKey(field.getName())) {
                    // 跳过已分配的索引
                    while (fieldColumnIndexMap.containsValue(autoIndex)) {
                        autoIndex++;
                    }
                    fieldColumnIndexMap.put(field.getName(), autoIndex);
                    log.debug("字段 {} 使用自动索引: {}", field.getName(), autoIndex);
                    autoIndex++;
                }
            }

            log.info("字段列索引映射结果: {}", fieldColumnIndexMap);
        } catch (Exception e) {
            log.error("构建字段列索引映射失败", e);
        }
    }

    /**
     * 处理图片字段
     * <p>
     * 自动兼容两种模式：
     * 1. 优先使用 POI 提取的 Drawing 图片（支持范围匹配，处理横跨/遮挡/超出情况）
     * 2. 如果没有 Drawing 图片，保留 EasyExcel 读取的单元格值（可能是 Base64 文本）
     * </p>
     * <p>
     * 图片匹配策略：
     * - 精确匹配：图片左上角锚点在目标单元格
     * - 范围匹配：图片横跨多个单元格时，检查目标单元格是否在图片覆盖范围内
     * - 行优先：同一行内，列匹配优先级最高
     * </p>
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
                    log.warn("字段 {} 未找到列索引映射", field.getName());
                    continue;
                }

                field.setAccessible(true);
                Class<?> fieldType = field.getType();

                // 使用范围匹配获取覆盖目标单元格的图片
                // 这会处理：1.精确匹配 2.横跨多列 3.遮挡单元格 4.超出一行
                List<POIImageExtractor.ImageData> images =
                    imageExtractor.getImagesCoveringCell(currentSheetIndex, rowIndex, columnIndex);

                // 如果范围匹配没找到，尝试精确匹配（后备方案）
                if (images.isEmpty()) {
                    images = imageExtractor.getImages(currentSheetIndex, rowIndex, columnIndex);
                }

                if (images.isEmpty()) {
                    // 没有 Drawing 图片，检查 EasyExcel 是否已经读取了 Base64 文本
                    Object existingValue = field.get(data);
                    if (existingValue != null) {
                        if (existingValue instanceof String && isBase64Image((String) existingValue)) {
                            log.debug("字段 {} 使用单元格中的 Base64 图片数据", field.getName());
                        } else if (existingValue instanceof byte[] && ((byte[]) existingValue).length > 0) {
                            log.debug("字段 {} 使用单元格中的字节数组图片数据", field.getName());
                        }
                    }
                    // 保留 EasyExcel 读取的值（可能是 Base64 或 null）
                    continue;
                }

                // 有 Drawing 图片，使用它覆盖字段值
                if (fieldType == String.class) {
                    // String类型：存储第一张图片的Base64
                    POIImageExtractor.ImageData imageData = images.get(0);
                    field.set(data, imageData.toBase64());
                    log.debug("设置图片字段 {} (String) 从 Drawing 对象，Row=[{}-{}], Col=[{}-{}], 大小: {} bytes",
                            field.getName(), imageData.getRow(), imageData.getRow2(),
                            imageData.getCol(), imageData.getCol2(), imageData.getData().length);

                } else if (fieldType == byte[].class) {
                    // byte[]类型：存储原始字节数据
                    POIImageExtractor.ImageData imageData = images.get(0);
                    field.set(data, imageData.getData());
                    log.debug("设置图片字段 {} (byte[]) 从 Drawing 对象，Row=[{}-{}], Col=[{}-{}], 大小: {} bytes",
                            field.getName(), imageData.getRow(), imageData.getRow2(),
                            imageData.getCol(), imageData.getCol2(), imageData.getData().length);

                } else if (fieldType == List.class) {
                    // List类型：存储多张图片
                    Type genericType = field.getGenericType();
                    if (genericType instanceof ParameterizedType) {
                        ParameterizedType pt = (ParameterizedType) genericType;
                        Type[] actualTypes = pt.getActualTypeArguments();

                        if (actualTypes.length > 0 && actualTypes[0] == String.class) {
                            // List<String>类型
                            List<String> imageList = new ArrayList<>();
                            Set<String> addedImageKeys = new HashSet<>();

                            // 对于 List 类型，只收集主锚点（左上角）在目标列的图片
                            // 不使用范围匹配，避免把其他列的大图片误收集进来
                            // 例如：thumbnail 在 Col=[4-5]，imageList 在列 5
                            // 范围匹配会把 thumbnail 也包含进来（因为 5 在 [4,5] 范围内）
                            for (POIImageExtractor.ImageData imageData : images) {
                                // 只添加主锚点在当前列的图片
                                if (imageData.getCol() == columnIndex) {
                                    String key = imageData.getRow() + "_" + imageData.getCol();
                                    if (!addedImageKeys.contains(key)) {
                                        imageList.add(imageData.toBase64());
                                        addedImageKeys.add(key);
                                    }
                                }
                            }

                            // 如果精确匹配没找到，尝试用精确匹配 API
                            if (imageList.isEmpty()) {
                                List<POIImageExtractor.ImageData> exactImages =
                                    imageExtractor.getImages(currentSheetIndex, rowIndex, columnIndex);
                                for (POIImageExtractor.ImageData imageData : exactImages) {
                                    String key = imageData.getRow() + "_" + imageData.getCol();
                                    if (!addedImageKeys.contains(key)) {
                                        imageList.add(imageData.toBase64());
                                        addedImageKeys.add(key);
                                    }
                                }
                            }

                            field.set(data, imageList);
                            log.debug("设置图片字段 {} (List<String>) 从 Drawing 对象，共 {} 张图片",
                                    field.getName(), imageList.size());
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("处理图片字段失败", e);
        }
    }

    /**
     * 检查字符串是否是 Base64 图片数据
     */
    private boolean isBase64Image(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        // 检查是否以 data:image 开头（标准 Base64 图片格式）
        if (value.startsWith("data:image/")) {
            return true;
        }
        // 检查是否是纯 Base64 字符串（至少100个字符，且只包含 Base64 字符）
        if (value.length() > 100) {
            return value.matches("^[A-Za-z0-9+/=]+$");
        }
        return false;
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