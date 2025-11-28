package cn.allbs.excel.aop;


import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.FlattenList;
import cn.allbs.excel.annotation.ImportExcel;
import cn.allbs.excel.annotation.ImportProgress;
import cn.allbs.excel.annotation.NestedProperty;
import cn.allbs.excel.convert.ImageBytesConverter;
import cn.allbs.excel.convert.LocalDateStringConverter;
import cn.allbs.excel.convert.LocalDateTimeStringConverter;
import cn.allbs.excel.convert.NestedObjectConverter;
import cn.allbs.excel.listener.ImportProgressListener;
import cn.allbs.excel.listener.ProgressReadListener;
import cn.allbs.excel.handle.DrawingImageReadListener;
import cn.allbs.excel.handle.HybridImageReadListener;
import cn.allbs.excel.handle.ImageAwareReadListener;
import cn.allbs.excel.handle.ListAnalysisEventListener;
import cn.allbs.excel.handle.SimpleImageReadListener;
import cn.allbs.excel.listener.FlattenListReadListener;
import cn.allbs.excel.model.ExcelReadError;
import cn.allbs.excel.vo.ErrorMessage;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.core.MethodParameter;
import org.springframework.core.ResolvableType;
import org.springframework.ui.ModelMap;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.support.WebDataBinderFactory;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartRequest;

import javax.servlet.http.HttpServletRequest;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.List;

/**
 * @author ChenQi
 */
@Slf4j
public class ImportExcelArgumentResolver implements HandlerMethodArgumentResolver, ApplicationContextAware {

    private ApplicationContext applicationContext;

    @Override
    public boolean supportsParameter(MethodParameter parameter) {
        return parameter.hasParameterAnnotation(ImportExcel.class);
    }

    @Override
    @SneakyThrows(Exception.class)
    public Object resolveArgument(MethodParameter parameter, ModelAndViewContainer modelAndViewContainer,
                                  NativeWebRequest webRequest, WebDataBinderFactory webDataBinderFactory) {
        Class<?> parameterType = parameter.getParameterType();
        if (!parameterType.isAssignableFrom(List.class)) {
            throw new IllegalArgumentException(
                    "Excel upload request resolver error, @RequestExcel parameter is not List " + parameterType);
        }

        // 处理自定义 readListener
        ImportExcel importExcel = parameter.getParameterAnnotation(ImportExcel.class);
        assert importExcel != null;
        Class<? extends ListAnalysisEventListener<?>> readListenerClass = importExcel.readListener();

        // 获取目标类型
        Class<?> excelModelClass = ResolvableType.forMethodParameter(parameter).getGeneric(0).resolve();

        // 获取请求文件流
        HttpServletRequest request = webRequest.getNativeRequest(HttpServletRequest.class);
        assert request != null;
        InputStream inputStream;
        byte[] excelBytes = null;

        // 检查是否需要图片支持
        ListAnalysisEventListener<?> readListener;

        if (needsImageSupport(excelModelClass)) {
            // 根据配置选择图片读取模式
            ImportExcel.ImageReadMode imageMode = importExcel.imageReadMode();

            switch (imageMode) {
                case DRAWING:
                    // 使用混合模式：POI读取图片 + EasyExcel读取数据
                    log.debug("Drawing模式：使用HybridImageReadListener（POI+EasyExcel混合）");

                    // 需要预读文件内容
                    if (request instanceof MultipartRequest) {
                        MultipartFile file = ((MultipartRequest) request).getFile(importExcel.fileName());
                        assert file != null;
                        excelBytes = file.getBytes();
                    } else {
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        byte[] buffer = new byte[1024];
                        int len;
                        InputStream reqStream = request.getInputStream();
                        while ((len = reqStream.read(buffer)) > -1) {
                            baos.write(buffer, 0, len);
                        }
                        excelBytes = baos.toByteArray();
                    }
                    inputStream = new ByteArrayInputStream(excelBytes);
                    readListener = new HybridImageReadListener(excelBytes);
                    break;

                case BASE64:
                    log.debug("Base64模式：使用SimpleImageReadListener（读取Base64文本）");
                    readListener = new SimpleImageReadListener();
                    // 获取普通输入流
                    if (request instanceof MultipartRequest) {
                        MultipartFile file = ((MultipartRequest) request).getFile(importExcel.fileName());
                        assert file != null;
                        inputStream = file.getInputStream();
                    } else {
                        inputStream = request.getInputStream();
                    }
                    break;

                case AUTO:
                default:
                    // 自动模式：使用HybridImageReadListener自动兼容Drawing图片和Base64文本
                    log.debug("自动模式：使用HybridImageReadListener（自动兼容Drawing图片和Base64文本）");

                    // 需要预读文件内容用于POI提取图片
                    if (request instanceof MultipartRequest) {
                        MultipartFile file = ((MultipartRequest) request).getFile(importExcel.fileName());
                        assert file != null;
                        excelBytes = file.getBytes();
                    } else {
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        byte[] buffer = new byte[1024];
                        int len;
                        InputStream reqStream = request.getInputStream();
                        while ((len = reqStream.read(buffer)) > -1) {
                            baos.write(buffer, 0, len);
                        }
                        excelBytes = baos.toByteArray();
                    }
                    inputStream = new ByteArrayInputStream(excelBytes);
                    readListener = new HybridImageReadListener(excelBytes);
                    break;
            }
        } else {
            readListener = BeanUtils.instantiateClass(readListenerClass);
            // 获取普通输入流
            if (request instanceof MultipartRequest) {
                MultipartFile file = ((MultipartRequest) request).getFile(importExcel.fileName());
                assert file != null;
                inputStream = file.getInputStream();
            } else {
                inputStream = request.getInputStream();
            }
        }

        // 检查是否需要导入进度支持
        ImportProgress importProgress = parameter.getMethodAnnotation(ImportProgress.class);
        if (importProgress != null && importProgress.enabled()) {
            log.debug("检测到 @ImportProgress 注解，启用导入进度回调");
            // 包装 readListener 以支持进度回调
            readListener = wrapWithProgressListener(readListener, importProgress);
        }

        // 检查是否需要 FlattenList 聚合处理
        if (needsFlattenListSupport(excelModelClass)) {
            log.debug("检测到 @FlattenList 注解，使用 FlattenListReadListener");
            FlattenListReadListener flattenListener = new FlattenListReadListener(excelModelClass);

            // 如果需要进度支持，包装 FlattenListReadListener
            if (importProgress != null && importProgress.enabled()) {
                FlattenListReadListener wrappedListener = this.wrapFlattenListWithProgressUnchecked(
                        flattenListener, importProgress, excelModelClass);
                EasyExcel.read(inputStream, wrappedListener)
                        .registerConverter(LocalDateStringConverter.INSTANCE)
                        .registerConverter(LocalDateTimeStringConverter.INSTANCE)
                        .sheet().doRead();
                return wrappedListener.getResult();
            }

            EasyExcel.read(inputStream, flattenListener)
                    .registerConverter(LocalDateStringConverter.INSTANCE)
                    .registerConverter(LocalDateTimeStringConverter.INSTANCE)
                    .sheet().doRead();
            return flattenListener.getResult();
        }

        // 这里需要指定读用哪个 class 去读，然后读取第一个 sheet 文件流会自动关闭
        // 构建 ExcelReaderBuilder
        ExcelReaderBuilder readerBuilder = EasyExcel.read(inputStream, excelModelClass, readListener)
                .registerConverter(LocalDateStringConverter.INSTANCE)
                .registerConverter(LocalDateTimeStringConverter.INSTANCE);

        // 如果需要图片支持，注册图片相关的转换器
        if (needsImageSupport(excelModelClass)) {
            readerBuilder.registerConverter(new ImageBytesConverter());
        }

        // 如果有 @NestedProperty 注解，注册 NestedObjectConverter
        if (needsNestedPropertySupport(excelModelClass)) {
            log.debug("检测到 @NestedProperty 注解，注册 NestedObjectConverter");
            readerBuilder.registerConverter(new NestedObjectConverter());
        }

        readerBuilder.ignoreEmptyRow(importExcel.ignoreEmptyRow())
                .sheet().doRead();

        // 校验失败的数据处理 交给 BindResult
        WebDataBinder dataBinder = webDataBinderFactory.createBinder(webRequest, readListener.getErrors(), "excel");
        ModelMap model = modelAndViewContainer.getModel();
        model.put(BindingResult.MODEL_KEY_PREFIX + "excel", dataBinder.getBindingResult());

        // 将错误信息放入 request attribute，方便控制器获取
        if (!readListener.getErrors().isEmpty()) {
            request.setAttribute("excelErrors", readListener.getErrors());
            log.debug("Excel导入发现{}个错误行", readListener.getErrors().size());
        }

        return readListener.getList();
    }

    /**
     * 检查是否需要图片支持
     *
     * @param excelModelClass Excel模型类
     * @return 如果包含@ExcelImage注解的字段返回true
     */
    private boolean needsImageSupport(Class<?> excelModelClass) {
        if (excelModelClass == null) {
            return false;
        }

        Field[] fields = excelModelClass.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelImage.class)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 检查是否需要 FlattenList 聚合支持
     *
     * @param excelModelClass Excel模型类
     * @return 如果包含@FlattenList注解的字段返回true
     */
    private boolean needsFlattenListSupport(Class<?> excelModelClass) {
        if (excelModelClass == null) {
            return false;
        }

        Field[] fields = excelModelClass.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(FlattenList.class)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 检查是否需要 NestedProperty 支持
     *
     * @param excelModelClass Excel模型类
     * @return 如果包含@NestedProperty注解的字段返回true
     */
    private boolean needsNestedPropertySupport(Class<?> excelModelClass) {
        if (excelModelClass == null) {
            return false;
        }

        Field[] fields = excelModelClass.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(NestedProperty.class)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }

    /**
     * 包装 ReadListener 以支持进度回调
     *
     * @param originalListener 原始监听器
     * @param importProgress   导入进度注解
     * @return 包装后的监听器
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private ListAnalysisEventListener<?> wrapWithProgressListener(
            ListAnalysisEventListener<?> originalListener,
            ImportProgress importProgress) {
        try {
            // 从Spring容器获取进度监听器bean，支持依赖注入
            ImportProgressListener progressListener;
            try {
                progressListener = applicationContext.getBean(importProgress.listener());
                log.debug("成功从Spring容器获取进度监听器: {}", importProgress.listener().getSimpleName());
            } catch (Exception e) {
                // 如果无法从容器获取，则手动实例化（不推荐，因为无法注入依赖）
                log.warn("无法从Spring容器获取进度监听器，尝试手动实例化: {}", importProgress.listener().getSimpleName());
                progressListener = BeanUtils.instantiateClass(importProgress.listener());
            }

            // 使用 ProgressReadListener 包装原始监听器
            return new ProgressReadListenerWrapper(
                    progressListener,
                    originalListener,
                    importProgress.interval()
            );
        } catch (Exception e) {
            log.error("Failed to wrap listener with progress support", e);
            return originalListener;
        }
    }

    /**
     * 包装 FlattenListReadListener 以支持进度回调（使用原始类型避免泛型推断问题）
     *
     * @param flattenListener FlattenList监听器
     * @param importProgress  导入进度注解
     * @param targetClass     目标类型
     * @return 包装后的监听器
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private FlattenListReadListener wrapFlattenListWithProgressUnchecked(
            FlattenListReadListener flattenListener,
            ImportProgress importProgress,
            Class targetClass) {
        try {
            // 从Spring容器获取进度监听器bean，支持依赖注入
            ImportProgressListener progressListener;
            try {
                progressListener = applicationContext.getBean(importProgress.listener());
                log.debug("成功从Spring容器获取进度监听器: {}", importProgress.listener().getSimpleName());
            } catch (Exception e) {
                // 如果无法从容器获取，则手动实例化（不推荐，因为无法注入依赖）
                log.warn("无法从Spring容器获取进度监听器，尝试手动实例化: {}", importProgress.listener().getSimpleName());
                progressListener = BeanUtils.instantiateClass(importProgress.listener());
            }

            // 使用反射创建包装类（需要特殊处理 FlattenListReadListener）
            return new FlattenListProgressWrapper(
                    progressListener,
                    flattenListener,
                    importProgress.interval(),
                    targetClass
            );
        } catch (Exception e) {
            log.error("Failed to wrap FlattenListReadListener with progress support", e);
            return flattenListener;
        }
    }

    /**
     * ProgressReadListener 包装器，用于包装 ListAnalysisEventListener
     */
    private static class ProgressReadListenerWrapper<T> extends ListAnalysisEventListener<T> {
        private final ProgressReadListener<T> progressListener;
        private final ListAnalysisEventListener<T> delegate;

        public ProgressReadListenerWrapper(ImportProgressListener progressListener,
                                            ListAnalysisEventListener<T> delegate,
                                            int interval) {
            this.progressListener = new ProgressReadListener<>(progressListener, delegate, interval);
            this.delegate = delegate;
        }

        @Override
        public void invoke(T data, com.alibaba.excel.context.AnalysisContext context) {
            progressListener.invoke(data, context);
        }

        @Override
        public void invokeHead(java.util.Map<Integer, com.alibaba.excel.metadata.data.ReadCellData<?>> headMap,
                               com.alibaba.excel.context.AnalysisContext context) {
            progressListener.invokeHead(headMap, context);
        }

        @Override
        public void doAfterAllAnalysed(com.alibaba.excel.context.AnalysisContext context) {
            progressListener.doAfterAllAnalysed(context);
        }

        @Override
        public void onException(Exception exception, com.alibaba.excel.context.AnalysisContext context) throws Exception {
            progressListener.onException(exception, context);
        }

        @Override
        public java.util.List<T> getList() {
            return delegate.getList();
        }

        @Override
        public java.util.List<ErrorMessage> getErrors() {
            return delegate.getErrors();
        }
    }

    /**
     * FlattenListReadListener 进度包装器
     * 注意：FlattenListReadListener 使用 Map<Integer, String> 而不是泛型 T
     */
    @SuppressWarnings("unchecked")
    private static class FlattenListProgressWrapper<T> extends FlattenListReadListener<T> {
        private final ProgressReadListener<java.util.Map<Integer, String>> progressListener;
        private final FlattenListReadListener<T> delegate;

        public FlattenListProgressWrapper(ImportProgressListener progressListener,
                                           FlattenListReadListener<T> delegate,
                                           int interval,
                                           Class<T> targetClass) {
            super(targetClass);
            this.progressListener = new ProgressReadListener<>(progressListener, delegate, interval);
            this.delegate = delegate;
        }

        @Override
        public void invokeHeadMap(java.util.Map<Integer, String> headMap, com.alibaba.excel.context.AnalysisContext context) {
            super.invokeHeadMap(headMap, context);
            if (progressListener != null) {
                // 转换为 ReadCellData 格式以兼容 ProgressReadListener
                java.util.Map<Integer, com.alibaba.excel.metadata.data.ReadCellData<?>> cellDataMap = new java.util.HashMap<>();
                headMap.forEach((k, v) -> {
                    com.alibaba.excel.metadata.data.ReadCellData<String> cellData = new com.alibaba.excel.metadata.data.ReadCellData<>();
                    cellData.setStringValue(v);
                    cellDataMap.put(k, cellData);
                });
                progressListener.invokeHead(cellDataMap, context);
            }
        }

        @Override
        public void invoke(java.util.Map<Integer, String> data, com.alibaba.excel.context.AnalysisContext context) {
            if (progressListener != null) {
                progressListener.invoke(data, context);
            }
            delegate.invoke(data, context);
        }

        @Override
        public void doAfterAllAnalysed(com.alibaba.excel.context.AnalysisContext context) {
            delegate.doAfterAllAnalysed(context);
            if (progressListener != null) {
                progressListener.doAfterAllAnalysed(context);
            }
        }

        @Override
        public void onException(Exception exception, com.alibaba.excel.context.AnalysisContext context) throws Exception {
            if (progressListener != null) {
                progressListener.onException(exception, context);
            }
            delegate.onException(exception, context);
        }

        @Override
        public java.util.List<T> getResult() {
            return delegate.getResult();
        }
    }

}
