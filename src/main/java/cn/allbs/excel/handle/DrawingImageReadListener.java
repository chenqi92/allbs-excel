package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import cn.allbs.excel.vo.FieldError;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.metadata.holder.ReadSheetHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import javax.validation.ConstraintViolation;
import javax.validation.metadata.ConstraintDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;

/**
 * Drawing图片读取监听器
 * <p>
 * 支持直接读取Excel中的Drawing对象（嵌入的图片）
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class DrawingImageReadListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();
    private final List<ErrorMessage> errorMessageList = new ArrayList<>();
    private Long lineNum = 1L;

    /**
     * 存储图片数据的映射
     * key: "row_col" 格式的字符串
     * value: 图片数据列表（一个单元格可能有多张图片）
     */
    private Map<String, List<PictureData>> pictureMap = new HashMap<>();

    /**
     * 是否已经加载了图片
     */
    private boolean picturesLoaded = false;

    /**
     * 当前处理的Sheet
     */
    private Sheet currentSheet;

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

    @Override
    public void invoke(Object data, AnalysisContext context) {
        lineNum++;

        try {
            // 第一次调用时加载图片
            if (!picturesLoaded) {
                loadPictures(context);
                picturesLoaded = true;
            }

            // 处理图片字段（传递 context 以获取正确的列映射）
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
     * 加载Sheet中的所有图片
     */
    private void loadPictures(AnalysisContext context) {
        try {
            ReadSheetHolder readSheetHolder = context.readSheetHolder();

            // 尝试多种方式获取Sheet对象
            currentSheet = getSheetFromContext(readSheetHolder);

            if (currentSheet == null) {
                log.warn("无法获取Sheet对象，跳过图片加载");
                return;
            }

            // 根据Excel类型处理图片
            if (currentSheet instanceof XSSFSheet) {
                loadXSSFPictures((XSSFSheet) currentSheet);
            } else if (currentSheet instanceof HSSFSheet) {
                loadHSSFPictures((HSSFSheet) currentSheet);
            } else {
                log.warn("不支持的Sheet类型: {}", currentSheet.getClass().getName());
            }

            log.info("加载了 {} 个单元格的图片", pictureMap.size());
        } catch (Exception e) {
            log.error("加载图片失败", e);
        }
    }

    /**
     * 尝试多种方式获取Sheet对象
     */
    private Sheet getSheetFromContext(ReadSheetHolder readSheetHolder) {
        // 方法1：尝试通过反射获取cachedSheet字段
        try {
            Field cachedSheetField = readSheetHolder.getClass().getDeclaredField("cachedSheet");
            cachedSheetField.setAccessible(true);
            Object cachedSheet = cachedSheetField.get(readSheetHolder);
            if (cachedSheet instanceof Sheet) {
                log.debug("通过cachedSheet字段获取到Sheet对象");
                return (Sheet) cachedSheet;
            }
        } catch (Exception e) {
            log.debug("无法通过cachedSheet字段获取Sheet: {}", e.getMessage());
        }

        // 方法2：尝试获取sheet字段
        try {
            Field sheetField = readSheetHolder.getClass().getDeclaredField("sheet");
            sheetField.setAccessible(true);
            Object sheet = sheetField.get(readSheetHolder);
            if (sheet instanceof Sheet) {
                log.debug("通过sheet字段获取到Sheet对象");
                return (Sheet) sheet;
            }
        } catch (Exception e) {
            log.debug("无法通过sheet字段获取Sheet: {}", e.getMessage());
        }

        // 方法3：尝试遍历所有字段查找Sheet类型
        try {
            Field[] fields = readSheetHolder.getClass().getDeclaredFields();
            for (Field field : fields) {
                if (Sheet.class.isAssignableFrom(field.getType())) {
                    field.setAccessible(true);
                    Object value = field.get(readSheetHolder);
                    if (value != null) {
                        log.debug("通过字段 {} 获取到Sheet对象", field.getName());
                        return (Sheet) value;
                    }
                }
            }

            // 还要检查父类的字段
            Class<?> superClass = readSheetHolder.getClass().getSuperclass();
            while (superClass != null && !superClass.equals(Object.class)) {
                fields = superClass.getDeclaredFields();
                for (Field field : fields) {
                    if (Sheet.class.isAssignableFrom(field.getType())) {
                        field.setAccessible(true);
                        Object value = field.get(readSheetHolder);
                        if (value != null) {
                            log.debug("通过父类字段 {} 获取到Sheet对象", field.getName());
                            return (Sheet) value;
                        }
                    }
                }
                superClass = superClass.getSuperclass();
            }
        } catch (Exception e) {
            log.debug("遍历字段查找Sheet失败: {}", e.getMessage());
        }

        return null;
    }

    /**
     * 加载XLSX格式的图片
     */
    private void loadXSSFPictures(XSSFSheet sheet) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            log.debug("Sheet中没有Drawing对象");
            return;
        }

        List<XSSFShape> shapes = drawing.getShapes();
        for (XSSFShape shape : shapes) {
            if (shape instanceof XSSFPicture) {
                XSSFPicture picture = (XSSFPicture) shape;
                XSSFClientAnchor anchor = picture.getClientAnchor();
                if (anchor != null) {
                    int row = anchor.getRow1();
                    int col = anchor.getCol1();
                    int row2 = anchor.getRow2();
                    int col2 = anchor.getCol2();
                    String key = row + "_" + col;

                    PictureData pictureData = new PictureData();
                    pictureData.data = picture.getPictureData().getData();
                    pictureData.mimeType = picture.getPictureData().getMimeType();
                    pictureData.pictureType = picture.getPictureData().getPictureType();
                    pictureData.row = row;
                    pictureData.col = col;
                    pictureData.row2 = row2;
                    pictureData.col2 = col2;

                    pictureMap.computeIfAbsent(key, k -> new ArrayList<>()).add(pictureData);

                    log.debug("找到图片: row=[{}-{}], col=[{}-{}], size={} bytes",
                              row, row2, col, col2, pictureData.data.length);
                }
            }
        }
    }

    /**
     * 加载XLS格式的图片
     */
    private void loadHSSFPictures(HSSFSheet sheet) {
        HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
        if (patriarch == null) {
            log.debug("Sheet中没有Drawing对象");
            return;
        }

        List<HSSFShape> shapes = patriarch.getChildren();
        for (HSSFShape shape : shapes) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor anchor = (HSSFClientAnchor) picture.getAnchor();
                if (anchor != null) {
                    int row = anchor.getRow1();
                    int col = anchor.getCol1();
                    int row2 = anchor.getRow2();
                    int col2 = anchor.getCol2();
                    String key = row + "_" + col;

                    PictureData pictureData = new PictureData();
                    pictureData.data = picture.getPictureData().getData();
                    pictureData.mimeType = picture.getPictureData().getMimeType();
                    pictureData.pictureType = picture.getPictureData().getFormat();
                    pictureData.row = row;
                    pictureData.col = col;
                    pictureData.row2 = row2;
                    pictureData.col2 = col2;

                    pictureMap.computeIfAbsent(key, k -> new ArrayList<>()).add(pictureData);

                    log.debug("找到图片: row=[{}-{}], col=[{}-{}], size={} bytes",
                              row, row2, col, col2, pictureData.data.length);
                }
            }
        }
    }

    /**
     * 获取覆盖指定单元格的图片（范围匹配）
     * <p>
     * 支持图片横跨多个单元格、遮挡单元格、超出行等情况
     * </p>
     */
    private List<PictureData> getPicturesCoveringCell(int targetRow, int targetCol) {
        List<PictureData> result = new ArrayList<>();

        // 遍历所有图片，检查是否覆盖目标单元格
        for (List<PictureData> pictures : pictureMap.values()) {
            for (PictureData picture : pictures) {
                if (picture.coversCell(targetRow, targetCol)) {
                    result.add(picture);
                }
            }
        }

        // 按主锚点位置排序（左上角优先）
        result.sort((a, b) -> {
            if (a.row != b.row) return a.row - b.row;
            return a.col - b.col;
        });

        return result;
    }

    /**
     * 处理图片字段
     * <p>
     * 支持范围匹配，处理图片横跨/遮挡/超出情况
     * </p>
     */
    private void processImageFields(Object data, AnalysisContext context) {
        try {
            Class<?> clazz = data.getClass();
            Field[] fields = clazz.getDeclaredFields();

            // 获取当前行索引
            // POI 中 Row 索引从0开始，表头在 Row=0，第一条数据在 Row=1
            // lineNum 初始为1，第一次 invoke 后 lineNum=2，所以 rowIndex = lineNum - 1 = 1
            int rowIndex = (int) (lineNum - 1);

            // 使用 EasyExcel 的 context 获取正确的字段到列索引的映射
            Map<Field, Integer> fieldColumnMap = getFieldColumnMap(clazz, context);

            for (Field field : fields) {
                if (!field.isAnnotationPresent(ExcelImage.class)) {
                    continue;
                }

                Integer columnIndex = fieldColumnMap.get(field);
                if (columnIndex == null) {
                    log.warn("字段 {} 未找到列索引映射", field.getName());
                    continue;
                }

                // 使用范围匹配获取覆盖目标单元格的图片
                List<PictureData> pictures = getPicturesCoveringCell(rowIndex, columnIndex);

                // 如果范围匹配没找到，尝试精确匹配（后备方案）
                if (pictures.isEmpty()) {
                    String key = rowIndex + "_" + columnIndex;
                    pictures = pictureMap.get(key);
                }

                if (pictures == null || pictures.isEmpty()) {
                    continue;
                }

                field.setAccessible(true);
                Class<?> fieldType = field.getType();

                if (fieldType == String.class) {
                    // String类型：存储第一张图片的Base64
                    PictureData pic = pictures.get(0);
                    String base64Data = "data:" + pic.mimeType + ";base64," +
                                       Base64.getEncoder().encodeToString(pic.data);
                    field.set(data, base64Data);
                    log.debug("设置图片字段 {} (String)", field.getName());

                } else if (fieldType == byte[].class) {
                    // byte[]类型：存储原始字节数据
                    PictureData pic = pictures.get(0);
                    field.set(data, pic.data);
                    log.debug("设置图片字段 {} (byte[])", field.getName());

                } else if (fieldType == List.class) {
                    // List类型：存储多张图片
                    Type genericType = field.getGenericType();
                    if (genericType instanceof ParameterizedType) {
                        ParameterizedType pt = (ParameterizedType) genericType;
                        Type[] actualTypes = pt.getActualTypeArguments();

                        if (actualTypes.length > 0 && actualTypes[0] == String.class) {
                            // List<String>类型
                            List<String> imageList = new ArrayList<>();
                            for (PictureData pic : pictures) {
                                String base64Data = "data:" + pic.mimeType + ";base64," +
                                                   Base64.getEncoder().encodeToString(pic.data);
                                imageList.add(base64Data);
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

    /**
     * 从 EasyExcel 的 AnalysisContext 获取字段到列索引的映射
     * <p>
     * 使用 invokeHeadMap 回调缓存的表头映射（headNameMap）
     * 这是最可靠的方式，因为 EasyExcel 在解析表头后会自动调用 invokeHeadMap
     * </p>
     */
    private Map<Field, Integer> getFieldColumnMap(Class<?> clazz, AnalysisContext context) {
        Map<Field, Integer> map = new HashMap<>();

        // headNameMap 已在 invokeHeadMap 回调中设置
        if (headNameMap.isEmpty()) {
            log.warn("表头映射为空，可能 invokeHeadMap 尚未被调用");
        }

        // 获取类的所有字段
        Field[] fields = clazz.getDeclaredFields();

        // 建立表头名称到字段的映射
        Map<String, Field> headNameToField = new HashMap<>();
        Map<Field, Integer> explicitIndexFields = new HashMap<>();

        for (Field field : fields) {
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                // 如果有显式指定 index，直接使用
                if (excelProperty.index() >= 0) {
                    explicitIndexFields.put(field, excelProperty.index());
                    map.put(field, excelProperty.index());
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
        if (!headNameMap.isEmpty()) {
            for (Map.Entry<Integer, String> entry : headNameMap.entrySet()) {
                Integer colIndex = entry.getKey();
                String headName = entry.getValue();

                Field matchedField = headNameToField.get(headName);
                if (matchedField != null && !map.containsKey(matchedField)) {
                    map.put(matchedField, colIndex);
                }
            }
        }

        // 对于仍未映射的字段，按声明顺序分配（作为后备方案）
        int autoIndex = 0;
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelProperty.class) && !map.containsKey(field)) {
                // 跳过已分配的索引
                while (map.containsValue(autoIndex)) {
                    autoIndex++;
                }
                map.put(field, autoIndex);
                autoIndex++;
            }
        }

        return map;
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed, total {} records with drawing images processed", list.size());
        // 清理资源
        pictureMap.clear();
        picturesLoaded = false;
        currentSheet = null;
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

    /**
     * 图片数据内部类
     */
    private static class PictureData {
        byte[] data;
        String mimeType;
        int pictureType;
        int row;      // 左上角行
        int col;      // 左上角列
        int row2;     // 右下角行
        int col2;     // 右下角列

        /**
         * 检查图片是否覆盖指定单元格
         */
        boolean coversCell(int targetRow, int targetCol) {
            // 精确匹配左上角
            if (row == targetRow && col == targetCol) {
                return true;
            }
            // 范围覆盖检查
            boolean rowInRange = row <= targetRow && targetRow <= row2;
            boolean colInRange = col <= targetCol && targetCol <= col2;
            return rowInRange && colInRange;
        }
    }
}