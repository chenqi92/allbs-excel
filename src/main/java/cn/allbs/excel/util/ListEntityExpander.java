package cn.allbs.excel.util;

import cn.allbs.excel.annotation.FlattenList;
import cn.allbs.excel.annotation.FlattenProperty;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;
import java.util.stream.Collectors;

/**
 * List 实体展开处理器
 * <p>
 * 将包含 List 字段的数据展开为多行
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class ListEntityExpander {

    /**
     * 展开数据列表
     *
     * @param dataList 原始数据列表
     * @param <T>      数据类型
     * @return 展开后的数据列表（使用 Map 表示）
     */
    public static <T> List<Map<String, Object>> expandData(List<T> dataList) {
        if (dataList == null || dataList.isEmpty()) {
            return Collections.emptyList();
        }

        Class<?> dataClass = dataList.get(0).getClass();
        ListExpandMetadata metadata = analyzeClass(dataClass);

        List<Map<String, Object>> result = new ArrayList<>();

        for (T data : dataList) {
            List<Map<String, Object>> expandedRows = expandSingleObject(data, metadata);
            result.addAll(expandedRows);
        }

        return result;
    }

    /**
     * 分析类结构，提取 List 字段和普通字段信息（带缓存）
     *
     * @param clazz 数据类
     * @return 展开元数据
     */
    public static ListExpandMetadata analyzeClass(Class<?> clazz) {
        return ClassMetadataCache.getOrCompute(
            ClassMetadataCache.CacheType.LIST_EXPAND,
            clazz,
            ListEntityExpander::doAnalyzeClass
        );
    }

    /**
     * 实际执行分析（内部方法）
     */
    private static ListExpandMetadata doAnalyzeClass(Class<?> clazz) {
        ListExpandMetadata metadata = new ListExpandMetadata();
        metadata.setDataClass(clazz);

        Field[] fields = clazz.getDeclaredFields();

        for (Field field : fields) {
            if (field.isAnnotationPresent(FlattenList.class)) {
                // List 字段
                FlattenList annotation = field.getAnnotation(FlattenList.class);
                ListFieldInfo listFieldInfo = new ListFieldInfo();
                listFieldInfo.setField(field);
                listFieldInfo.setAnnotation(annotation);

                // 获取泛型类型
                Type genericType = field.getGenericType();
                if (genericType instanceof ParameterizedType) {
                    ParameterizedType paramType = (ParameterizedType) genericType;
                    Type[] actualTypes = paramType.getActualTypeArguments();
                    if (actualTypes.length > 0 && actualTypes[0] instanceof Class) {
                        listFieldInfo.setElementType((Class<?>) actualTypes[0]);

                        // 扫描元素类型的字段
                        List<ExpandedField> elementFields = scanElementFields((Class<?>) actualTypes[0], annotation);
                        listFieldInfo.setElementFields(elementFields);
                    }
                }

                metadata.getListFields().add(listFieldInfo);
            } else if (field.isAnnotationPresent(FlattenProperty.class)) {
                // @FlattenProperty 字段 - 需要展开其子字段作为普通字段
                FlattenProperty flattenProperty = field.getAnnotation(FlattenProperty.class);
                Class<?> fieldType = field.getType();

                // 扫描嵌套对象的字段
                Field[] nestedFields = fieldType.getDeclaredFields();
                for (Field nestedField : nestedFields) {
                    if (nestedField.isAnnotationPresent(ExcelProperty.class)) {
                        ExcelProperty excelProperty = nestedField.getAnnotation(ExcelProperty.class);
                        String[] headNames = excelProperty.value();
                        String headName = headNames.length > 0 ? headNames[0] : nestedField.getName();

                        // 应用前缀和后缀
                        String finalHeadName = flattenProperty.prefix() + headName + flattenProperty.suffix();

                        // 创建 NormalFieldInfo，但需要记录它是从 FlattenProperty 展开的
                        FlattenPropertyFieldInfo flattenFieldInfo = new FlattenPropertyFieldInfo();
                        flattenFieldInfo.setField(nestedField);
                        flattenFieldInfo.setParentField(field);
                        flattenFieldInfo.setExcelProperty(excelProperty);
                        flattenFieldInfo.setFinalHeadName(finalHeadName);
                        metadata.getNormalFields().add(flattenFieldInfo);
                    }
                }
            } else if (field.isAnnotationPresent(ExcelProperty.class)) {
                // 普通字段
                NormalFieldInfo normalFieldInfo = new NormalFieldInfo();
                normalFieldInfo.setField(field);
                normalFieldInfo.setExcelProperty(field.getAnnotation(ExcelProperty.class));
                metadata.getNormalFields().add(normalFieldInfo);
            }
        }

        // 按顺序排序 List 字段
        metadata.getListFields().sort(Comparator.comparingInt(f -> f.getAnnotation().order()));

        return metadata;
    }

    /**
     * 扫描 List 元素类型的字段
     *
     * @param elementClass 元素类型
     * @param annotation   FlattenList 注解
     * @return 展开字段列表
     */
    private static List<ExpandedField> scanElementFields(Class<?> elementClass, FlattenList annotation) {
        List<ExpandedField> result = new ArrayList<>();
        Field[] fields = elementClass.getDeclaredFields();

        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelProperty.class)) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                String[] headNames = excelProperty.value();
                String headName = headNames.length > 0 ? headNames[0] : field.getName();

                // 应用前缀和后缀
                String finalHeadName = annotation.prefix() + headName + annotation.suffix();

                ExpandedField expandedField = new ExpandedField();
                expandedField.setField(field);
                expandedField.setOriginalHeadName(headName);
                expandedField.setFinalHeadName(finalHeadName);
                expandedField.setExcelProperty(excelProperty);

                result.add(expandedField);
            }
        }

        return result;
    }

    /**
     * 展开单个对象
     *
     * @param obj      对象
     * @param metadata 元数据
     * @return 展开后的行列表
     */
    private static List<Map<String, Object>> expandSingleObject(Object obj, ListExpandMetadata metadata) {
        List<Map<String, Object>> result = new ArrayList<>();

        try {
            // 获取所有 List 字段的值
            List<ListFieldValue> listFieldValues = new ArrayList<>();
            int maxListSize = 0;

            for (ListFieldInfo listFieldInfo : metadata.getListFields()) {
                Object listValue = getFieldValue(obj, listFieldInfo.getField());
                List<?> list = (listValue instanceof List) ? (List<?>) listValue : Collections.emptyList();

                // 处理 maxRows 限制
                if (listFieldInfo.getAnnotation().maxRows() > 0 && list.size() > listFieldInfo.getAnnotation().maxRows()) {
                    list = list.subList(0, listFieldInfo.getAnnotation().maxRows());
                }

                ListFieldValue fieldValue = new ListFieldValue();
                fieldValue.setListFieldInfo(listFieldInfo);
                fieldValue.setList(list);
                listFieldValues.add(fieldValue);

                maxListSize = Math.max(maxListSize, list.size());
            }

            // 确定展开行数（根据策略）
            int expandedRowCount = calculateExpandedRowCount(listFieldValues, metadata);

            // 如果没有 List 或全部为空，至少保留一行
            if (expandedRowCount == 0) {
                expandedRowCount = 1;
            }

            // 生成展开后的行
            for (int rowIndex = 0; rowIndex < expandedRowCount; rowIndex++) {
                Map<String, Object> row = new LinkedHashMap<>();

                // 添加普通字段（每行都相同）
                for (NormalFieldInfo normalField : metadata.getNormalFields()) {
                    Object value;
                    if (normalField instanceof FlattenPropertyFieldInfo) {
                        // FlattenProperty 字段：需要先获取父对象，再获取子字段值
                        FlattenPropertyFieldInfo flattenField = (FlattenPropertyFieldInfo) normalField;
                        Object parentObj = getFieldValue(obj, flattenField.getParentField());
                        if (parentObj != null) {
                            value = getFieldValue(parentObj, flattenField.getField());
                        } else {
                            value = null;
                        }
                    } else {
                        // 普通字段：直接获取值
                        value = getFieldValue(obj, normalField.getField());
                    }
                    row.put(normalField.getHeadName(), value);
                }

                // 添加 List 展开字段
                for (ListFieldValue listFieldValue : listFieldValues) {
                    List<?> list = listFieldValue.getList();
                    ListFieldInfo listFieldInfo = listFieldValue.getListFieldInfo();

                    // 获取当前行对应的元素
                    Object element = null;
                    if (rowIndex < list.size()) {
                        element = list.get(rowIndex);
                    }

                    // 提取元素的字段值
                    for (ExpandedField expandedField : listFieldInfo.getElementFields()) {
                        Object fieldValue = null;
                        if (element != null) {
                            fieldValue = getFieldValue(element, expandedField.getField());
                        }
                        row.put(expandedField.getFinalHeadName(), fieldValue);
                    }
                }

                result.add(row);
            }

        } catch (Exception e) {
            log.error("Failed to expand object", e);
        }

        return result;
    }

    /**
     * 计算展开行数
     *
     * @param listFieldValues List 字段值列表
     * @param metadata        元数据
     * @return 展开行数
     */
    private static int calculateExpandedRowCount(List<ListFieldValue> listFieldValues, ListExpandMetadata metadata) {
        if (listFieldValues.isEmpty()) {
            return 1;
        }

        // 获取第一个 List 的策略（假设所有 List 使用相同策略）
        FlattenList.MultiListStrategy strategy = listFieldValues.get(0).getListFieldInfo()
                .getAnnotation().multiListStrategy();

        switch (strategy) {
            case MAX_LENGTH:
                return listFieldValues.stream()
                        .mapToInt(v -> v.getList().size())
                        .max()
                        .orElse(1);

            case MIN_LENGTH:
                return listFieldValues.stream()
                        .mapToInt(v -> v.getList().size())
                        .filter(size -> size > 0)
                        .min()
                        .orElse(1);

            case CARTESIAN:
                // 笛卡尔积：所有 List 大小的乘积
                return listFieldValues.stream()
                        .mapToInt(v -> Math.max(v.getList().size(), 1))
                        .reduce(1, (a, b) -> a * b);

            default:
                return 1;
        }
    }

    /**
     * 获取字段值
     *
     * @param obj   对象
     * @param field 字段
     * @return 字段值
     */
    private static Object getFieldValue(Object obj, Field field) {
        try {
            // 优先使用 getter 方法
            String getterName = "get" + capitalize(field.getName());
            try {
                Method getter = obj.getClass().getMethod(getterName);
                return getter.invoke(obj);
            } catch (NoSuchMethodException e) {
                // 尝试 is 开头的 getter
                try {
                    String isGetterName = "is" + capitalize(field.getName());
                    Method isGetter = obj.getClass().getMethod(isGetterName);
                    return isGetter.invoke(obj);
                } catch (NoSuchMethodException ex) {
                    // 直接访问字段
                    field.setAccessible(true);
                    return field.get(obj);
                }
            }
        } catch (Exception e) {
            log.error("Failed to get field value: {}", field.getName(), e);
            return null;
        }
    }

    /**
     * 首字母大写
     */
    private static String capitalize(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }

    /**
     * 生成合并单元格配置
     *
     * @param expandedData  展开后的数据
     * @param metadata      元数据
     * @return 合并区域列表
     */
    public static List<MergeRegion> generateMergeRegions(List<Map<String, Object>> expandedData,
                                                          ListExpandMetadata metadata) {
        List<MergeRegion> regions = new ArrayList<>();

        log.debug("Generating merge regions for {} expanded rows", expandedData.size());
        log.debug("Normal fields count: {}", metadata.getNormalFields().size());
        for (NormalFieldInfo field : metadata.getNormalFields()) {
            log.debug("Normal field: {}", field.getHeadName());
        }

        // 生成表头列名列表
        List<String> headers = generateHeaders(metadata);

        int currentRow = 1; // Excel 行从 1 开始（0 是表头）
        int dataIndex = 0;

        while (dataIndex < expandedData.size()) {
            // 找到连续的相同记录（展开前是同一条记录）
            int groupSize = 1;
            Map<String, Object> firstRow = expandedData.get(dataIndex);

            // 通过普通字段判断是否是同一条原始记录
            for (int i = dataIndex + 1; i < expandedData.size(); i++) {
                Map<String, Object> currentRowData = expandedData.get(i);
                boolean isSameGroup = true;

                for (NormalFieldInfo normalField : metadata.getNormalFields()) {
                    String key = normalField.getHeadName();

                    Object val1 = firstRow.get(key);
                    Object val2 = currentRowData.get(key);

                    if (!Objects.equals(val1, val2)) {
                        isSameGroup = false;
                        break;
                    }
                }

                if (isSameGroup) {
                    groupSize++;
                } else {
                    break;
                }
            }

            // 如果 groupSize > 1，需要合并普通字段的单元格
            if (groupSize > 1) {
                log.debug("Found group with size {} at row {}", groupSize, currentRow);
                for (NormalFieldInfo normalField : metadata.getNormalFields()) {
                    String key = normalField.getHeadName();

                    int columnIndex = headers.indexOf(key);
                    if (columnIndex >= 0) {
                        MergeRegion region = new MergeRegion();
                        region.setFirstRow(currentRow);
                        region.setLastRow(currentRow + groupSize - 1);
                        region.setFirstColumn(columnIndex);
                        region.setLastColumn(columnIndex);
                        regions.add(region);
                        log.debug("Created merge region for column '{}' (index {}): rows {}-{}",
                                 key, columnIndex, currentRow, currentRow + groupSize - 1);
                    }
                }
            }

            currentRow += groupSize;
            dataIndex += groupSize;
        }

        return regions;
    }

    /**
     * 生成表头列名列表
     *
     * @param metadata 元数据
     * @return 表头列表
     */
    public static List<String> generateHeaders(ListExpandMetadata metadata) {
        List<String> headers = new ArrayList<>();

        // 添加普通字段表头（包括 FlattenProperty 展开的字段）
        for (NormalFieldInfo normalField : metadata.getNormalFields()) {
            headers.add(normalField.getHeadName());
        }

        // 添加 List 展开字段表头
        for (ListFieldInfo listField : metadata.getListFields()) {
            for (ExpandedField expandedField : listField.getElementFields()) {
                headers.add(expandedField.getFinalHeadName());
            }
        }

        return headers;
    }

    /**
     * 将 Map 数据转换为 List 数据（按表头顺序）
     * <p>
     * EasyExcel 在写入 List&lt;Map&gt; 时不会按表头顺序写入，需要转换为 List&lt;List&gt; 格式
     * </p>
     *
     * @param expandedData 展开后的 Map 数据
     * @param headers      表头列表
     * @return List 格式的数据
     */
    public static List<List<Object>> convertToListData(List<Map<String, Object>> expandedData,
                                                        List<String> headers) {
        List<List<Object>> result = new ArrayList<>();

        for (Map<String, Object> rowMap : expandedData) {
            List<Object> row = new ArrayList<>();
            for (String header : headers) {
                row.add(rowMap.get(header));
            }
            result.add(row);
        }

        return result;
    }

    // ==================== 数据类 ====================

    /**
     * List 展开元数据
     */
    @Data
    public static class ListExpandMetadata {
        private Class<?> dataClass;
        private List<NormalFieldInfo> normalFields = new ArrayList<>();
        private List<ListFieldInfo> listFields = new ArrayList<>();
    }

    /**
     * 普通字段信息
     */
    @Data
    public static class NormalFieldInfo {
        private Field field;
        private ExcelProperty excelProperty;

        /**
         * 获取表头名称
         */
        public String getHeadName() {
            String[] headNames = excelProperty.value();
            return headNames.length > 0 ? headNames[0] : field.getName();
        }
    }

    /**
     * FlattenProperty 展开字段信息
     */
    @Data
    public static class FlattenPropertyFieldInfo extends NormalFieldInfo {
        private Field parentField;  // 父字段（@FlattenProperty 字段）
        private String finalHeadName;  // 应用前缀后缀后的最终表头名称

        @Override
        public String getHeadName() {
            return finalHeadName;
        }
    }

    /**
     * List 字段信息
     */
    @Data
    public static class ListFieldInfo {
        private Field field;
        private FlattenList annotation;
        private Class<?> elementType;
        private List<ExpandedField> elementFields;
    }

    /**
     * 展开字段
     */
    @Data
    public static class ExpandedField {
        private Field field;
        private String originalHeadName;
        private String finalHeadName;
        private ExcelProperty excelProperty;
    }

    /**
     * List 字段值
     */
    @Data
    private static class ListFieldValue {
        private ListFieldInfo listFieldInfo;
        private List<?> list;
    }

    /**
     * 合并区域
     */
    @Data
    public static class MergeRegion {
        private int firstRow;
        private int lastRow;
        private int firstColumn;
        private int lastColumn;
    }
}
