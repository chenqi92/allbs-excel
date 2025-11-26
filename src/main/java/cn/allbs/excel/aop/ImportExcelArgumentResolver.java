package cn.allbs.excel.aop;


import cn.allbs.excel.annotation.ExcelImage;
import cn.allbs.excel.annotation.FlattenList;
import cn.allbs.excel.annotation.ImportExcel;
import cn.allbs.excel.annotation.NestedProperty;
import cn.allbs.excel.convert.ImageBytesConverter;
import cn.allbs.excel.convert.LocalDateStringConverter;
import cn.allbs.excel.convert.LocalDateTimeStringConverter;
import cn.allbs.excel.convert.NestedObjectConverter;
import cn.allbs.excel.handle.DrawingImageReadListener;
import cn.allbs.excel.handle.HybridImageReadListener;
import cn.allbs.excel.handle.ImageAwareReadListener;
import cn.allbs.excel.handle.ListAnalysisEventListener;
import cn.allbs.excel.handle.SimpleImageReadListener;
import cn.allbs.excel.listener.FlattenListReadListener;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.BeanUtils;
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
public class ImportExcelArgumentResolver implements HandlerMethodArgumentResolver {

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

        // 检查是否需要 FlattenList 聚合处理
        if (needsFlattenListSupport(excelModelClass)) {
            log.debug("检测到 @FlattenList 注解，使用 FlattenListReadListener");
            FlattenListReadListener<?> flattenListener = new FlattenListReadListener<>(excelModelClass);
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

}
