package cn.allbs.excel.handle;

import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.annotation.ExportProgress;
import cn.allbs.excel.annotation.Sheet;
import cn.allbs.excel.annotation.ExcelWatermark;
import cn.allbs.excel.annotation.ExcelChart;
import cn.allbs.excel.aop.DynamicNameAspect;
import cn.allbs.excel.aop.ExportExcelReturnValueHandler;
import cn.allbs.excel.convert.LocalDateStringConverter;
import cn.allbs.excel.convert.LocalDateTimeStringConverter;
import cn.allbs.excel.convert.NestedObjectConverter;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import cn.allbs.excel.head.HeadGenerator;
import cn.allbs.excel.head.HeadMeta;
import cn.allbs.excel.head.I18nHeaderCellWriteHandler;
import cn.allbs.excel.listener.ExportProgressListener;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.MediaTypeFactory;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Modifier;
import java.net.URLEncoder;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.UUID;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
@Slf4j
@RequiredArgsConstructor
public abstract class AbstractSheetWriteHandler implements SheetWriteHandler, ApplicationContextAware {

    private final ExcelConfigProperties configProperties;

    private final ObjectProvider<List<Converter<?>>> converterProvider;

    private final WriterBuilderEnhancer excelWriterBuilderEnhance;

    private ApplicationContext applicationContext;

    @Getter
    @Setter
    @Autowired(required = false)
    private I18nHeaderCellWriteHandler i18nHeaderCellWriteHandler;

    @Override
    public void check(ExportExcel responseExcel) {
        if (responseExcel.sheets().length == 0) {
            throw new ExcelException("@ResponseExcel sheet 配置不合法");
        }
    }

    @Override
    @SneakyThrows(UnsupportedEncodingException.class)
    public void export(Object o, HttpServletResponse response, ExportExcel responseExcel) {
        check(responseExcel);
        RequestAttributes requestAttributes = RequestContextHolder.getRequestAttributes();
        String name = (String) Objects.requireNonNull(requestAttributes).getAttribute(DynamicNameAspect.EXCEL_NAME_KEY,
                RequestAttributes.SCOPE_REQUEST);
        if (name == null) {
            name = UUID.randomUUID().toString();
        }
        String fileName = String.format("%s%s", URLEncoder.encode(name, "UTF-8"), responseExcel.suffix().getValue());
        // 根据实际的文件类型找到对应的 contentType
        String contentType = MediaTypeFactory.getMediaType(fileName).map(MediaType::toString)
                .orElse("application/vnd.ms-excel");
        response.setContentType(contentType);
        response.setCharacterEncoding("utf-8");
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename*=utf-8''" + fileName);
        response.setHeader(HttpHeaders.ACCESS_CONTROL_EXPOSE_HEADERS, HttpHeaders.CONTENT_DISPOSITION);
        write(o, response, responseExcel);
    }

    /**
     * 通用的获取ExcelWriter方法
     *
     * @param response      HttpServletResponse
     * @param responseExcel ResponseExcel注解
     * @return ExcelWriter
     */
    @SneakyThrows(IOException.class)
    public ExcelWriter getExcelWriter(HttpServletResponse response, ExportExcel responseExcel) {
        ExcelWriterBuilder writerBuilder = EasyExcel.write(response.getOutputStream())
                .registerConverter(LocalDateStringConverter.INSTANCE)
                .registerConverter(LocalDateTimeStringConverter.INSTANCE)
                .registerConverter(new NestedObjectConverter())
                .autoCloseStream(true)
                .excelType(responseExcel.suffix()).inMemory(responseExcel.inMemory());

        if (StringUtils.hasText(responseExcel.password())) {
            writerBuilder.password(responseExcel.password());
        }

        if (responseExcel.include().length != 0) {
            writerBuilder.includeColumnFieldNames(Arrays.asList(responseExcel.include()));
        }

        if (responseExcel.exclude().length != 0) {
            writerBuilder.excludeColumnFieldNames(Arrays.asList(responseExcel.exclude()));
        }

        if (responseExcel.writeHandler().length != 0) {
            for (Class<? extends WriteHandler> clazz : responseExcel.writeHandler()) {
                writerBuilder.registerWriteHandler(BeanUtils.instantiateClass(clazz));
            }
        }

        // 开启国际化头信息处理
        if (responseExcel.i18nHeader() && i18nHeaderCellWriteHandler != null) {
            writerBuilder.registerWriteHandler(i18nHeaderCellWriteHandler);
        }

        // 处理水印配置
        ExcelWatermark watermark = responseExcel.watermark();
        if (watermark != null && watermark.enabled() && StringUtils.hasText(watermark.text())) {
            writerBuilder.registerWriteHandler(new cn.allbs.excel.handle.ExcelWatermarkWriteHandler(watermark));
        }

        // 自定义注入的转换器
        registerCustomConverter(writerBuilder);

        for (Class<? extends Converter> clazz : responseExcel.converter()) {
            writerBuilder.registerConverter(BeanUtils.instantiateClass(clazz));
        }

        String templatePath = configProperties.getTemplatePath();
        if (StringUtils.hasText(responseExcel.template())) {
            ClassPathResource classPathResource = new ClassPathResource(
                    templatePath + File.separator + responseExcel.template());
            InputStream inputStream = classPathResource.getInputStream();
            writerBuilder.withTemplate(inputStream);
        }

        writerBuilder = excelWriterBuilderEnhance.enhanceExcel(writerBuilder, response, responseExcel, templatePath);

        return writerBuilder.build();
    }

    /**
     * 自定义注入转换器 如果有需要，子类自己重写
     *
     * @param builder ExcelWriterBuilder
     */
    public void registerCustomConverter(ExcelWriterBuilder builder) {
        converterProvider.ifAvailable(converters -> converters.forEach(builder::registerConverter));
    }

    /**
     * 获取 WriteSheet 对象
     *
     * @param sheet                 sheet annotation info
     * @param dataClass             数据类型
     * @param template              模板
     * @param bookHeadEnhancerClass 自定义头处理器
     * @return WriteSheet
     */
    public WriteSheet sheet(Sheet sheet, Class<?> dataClass, String template,
                            Class<? extends HeadGenerator> bookHeadEnhancerClass) {
        return sheet(sheet, dataClass, template, bookHeadEnhancerClass, false);
    }

    /**
     * 获取 WriteSheet 对象
     *
     * @param sheet                 sheet annotation info
     * @param dataClass             数据类型
     * @param template              模板
     * @param bookHeadEnhancerClass 自定义头处理器
     * @param onlyExcelProperty     是否只导出有 @ExcelProperty 注解的字段
     * @return WriteSheet
     */
    public WriteSheet sheet(Sheet sheet, Class<?> dataClass, String template,
                            Class<? extends HeadGenerator> bookHeadEnhancerClass, boolean onlyExcelProperty) {
        return sheet(sheet, dataClass, template, bookHeadEnhancerClass, onlyExcelProperty, false);
    }

    /**
     * 获取 WriteSheet 对象
     *
     * @param sheet                 sheet annotation info
     * @param dataClass             数据类型
     * @param template              模板
     * @param bookHeadEnhancerClass 自定义头处理器
     * @param onlyExcelProperty     是否只导出有 @ExcelProperty 注解的字段
     * @param autoMerge             是否自动合并相同值的单元格
     * @return WriteSheet
     */
    public WriteSheet sheet(Sheet sheet, Class<?> dataClass, String template,
                            Class<? extends HeadGenerator> bookHeadEnhancerClass, boolean onlyExcelProperty,
                            boolean autoMerge) {
        return sheet(sheet, dataClass, template, bookHeadEnhancerClass, onlyExcelProperty, autoMerge, 0);
    }

    /**
     * 获取 WriteSheet 对象
     *
     * @param sheet                 sheet annotation info
     * @param dataClass             数据类型
     * @param template              模板
     * @param bookHeadEnhancerClass 自定义头处理器
     * @param onlyExcelProperty     是否只导出有 @ExcelProperty 注解的字段
     * @param autoMerge             是否自动合并相同值的单元格
     * @param totalRows             总行数（用于进度回调）
     * @return WriteSheet
     */
    public WriteSheet sheet(Sheet sheet, Class<?> dataClass, String template,
                            Class<? extends HeadGenerator> bookHeadEnhancerClass, boolean onlyExcelProperty,
                            boolean autoMerge, int totalRows) {

        // Sheet 编号和名称
        Integer sheetNo = sheet.sheetNo() >= 0 ? sheet.sheetNo() : null;
        String sheetName = sheet.sheetName();

        // 是否模板写入
        ExcelWriterSheetBuilder writerSheetBuilder = StringUtils.hasText(template) ? EasyExcel.writerSheet(sheetNo)
                : EasyExcel.writerSheet(sheetNo, sheetName);

        // 头信息增强 1. 优先使用 sheet 指定的头信息增强 2. 其次使用 @ResponseExcel 中定义的全局头信息增强
        Class<? extends HeadGenerator> headGenerateClass = null;
        if (isNotInterface(sheet.headGenerateClass())) {
            headGenerateClass = sheet.headGenerateClass();
        } else if (isNotInterface(bookHeadEnhancerClass)) {
            headGenerateClass = bookHeadEnhancerClass;
        }
        // 定义头信息增强则使用其生成头信息，否则使用 dataClass 来自动获取
        if (headGenerateClass != null) {
            fillCustomHeadInfo(dataClass, bookHeadEnhancerClass, writerSheetBuilder);
        } else if (dataClass != null) {
            writerSheetBuilder.head(dataClass);

            // 处理 onlyExcelProperty 配置
            // 优先级：Sheet 级别 > 全局级别
            boolean shouldOnlyExcelProperty = sheet.onlyExcelProperty() || onlyExcelProperty;
            if (shouldOnlyExcelProperty) {
                // 只导出有 @ExcelProperty 注解的字段
                writerSheetBuilder.excludeColumnFieldNames(getFieldsWithoutExcelProperty(dataClass));
            }

            if (sheet.excludes().length > 0) {
                writerSheetBuilder.excludeColumnFieldNames(Arrays.asList(sheet.excludes()));
            }
            if (sheet.includes().length > 0) {
                writerSheetBuilder.includeColumnFieldNames(Arrays.asList(sheet.includes()));
            }
        }

        // 处理 autoMerge 配置
        // 优先级：Sheet 级别 > 全局级别
        boolean shouldAutoMerge = sheet.autoMerge() || autoMerge;
        if (shouldAutoMerge && dataClass != null) {
            // 添加合并处理器
            MergeCellWriteHandler mergeCellWriteHandler = new MergeCellWriteHandler(dataClass);
            writerSheetBuilder.registerWriteHandler(mergeCellWriteHandler);
        }

        // 处理进度回调
        if (totalRows > 0) {
            RequestAttributes requestAttributes = RequestContextHolder.getRequestAttributes();
            if (requestAttributes != null) {
                ExportProgress exportProgress = (ExportProgress) requestAttributes.getAttribute(
                        ExportExcelReturnValueHandler.EXPORT_PROGRESS_KEY, RequestAttributes.SCOPE_REQUEST);
                if (exportProgress != null && exportProgress.enabled()) {
                    try {
                        // 从 Spring 容器中获取进度监听器实例，确保依赖注入生效
                        ExportProgressListener listener = applicationContext.getBean(exportProgress.listener());
                        // 创建进度处理器
                        ProgressWriteHandler progressWriteHandler = new ProgressWriteHandler(
                                listener, totalRows, exportProgress.interval(), sheetName);
                        writerSheetBuilder.registerWriteHandler(progressWriteHandler);
                    } catch (Exception e) {
                        throw new ExcelException("Failed to create progress listener: " + e.getMessage());
                    }
                }
            }
        }

        // 处理数据验证（ExcelValidation）
        if (dataClass != null && hasExcelValidationAnnotation(dataClass)) {
            int startRow = sheet.validationStartRow() >= 0 ? sheet.validationStartRow() : 1;
            int endRow = sheet.validationEndRow() >= 0 ? sheet.validationEndRow() : 65535;
            cn.allbs.excel.handle.ExcelValidationWriteHandler validationHandler =
                new cn.allbs.excel.handle.ExcelValidationWriteHandler(dataClass, startRow, endRow);
            writerSheetBuilder.registerWriteHandler(validationHandler);
        }

        // 处理条件样式（ConditionalStyle）
        if (dataClass != null && hasConditionalStyleAnnotation(dataClass)) {
            cn.allbs.excel.handle.ConditionalStyleWriteHandler conditionalStyleHandler =
                new cn.allbs.excel.handle.ConditionalStyleWriteHandler(dataClass);
            writerSheetBuilder.registerWriteHandler(conditionalStyleHandler);
            log.debug("Registered ConditionalStyleWriteHandler for dataClass: {}", dataClass.getSimpleName());
        }

        // 处理图片导出（ExcelImage）
        if (dataClass != null && hasExcelImageAnnotation(dataClass)) {
            cn.allbs.excel.handle.ImageWriteHandler imageHandler =
                new cn.allbs.excel.handle.ImageWriteHandler(dataClass);
            writerSheetBuilder.registerWriteHandler(imageHandler);
            log.debug("Registered ImageWriteHandler for dataClass: {}", dataClass.getSimpleName());
        }

        // 处理图表配置（ExcelChart） - 从 RequestAttributes 获取
        RequestAttributes requestAttributes = RequestContextHolder.getRequestAttributes();
        if (requestAttributes != null && dataClass != null && totalRows > 0) {
            ExcelChart chart = (ExcelChart) requestAttributes.getAttribute(
                    "EXPORT_CHART_CONFIG", RequestAttributes.SCOPE_REQUEST);
            log.debug("Retrieved chart config from request attributes: {}", chart);
            if (chart != null && chart.enabled() && StringUtils.hasText(chart.title())) {
                log.info("Registering ExcelChartWriteHandler: title={}, type={}, xAxis={}, yAxis={}",
                        chart.title(), chart.type(), chart.xAxisField(), java.util.Arrays.toString(chart.yAxisFields()));
                cn.allbs.excel.handle.ExcelChartWriteHandler chartHandler =
                    new cn.allbs.excel.handle.ExcelChartWriteHandler(chart, dataClass, 1, totalRows);
                writerSheetBuilder.registerWriteHandler(chartHandler);
            } else {
                log.debug("Chart not registered: chart={}, enabled={}, hasTitle={}",
                        chart != null, chart != null && chart.enabled(), chart != null && StringUtils.hasText(chart.title()));
            }
        }

        // sheetBuilder 增强
        writerSheetBuilder = excelWriterBuilderEnhance.enhanceSheet(writerSheetBuilder, sheetNo, sheetName, dataClass,
                template, headGenerateClass);

        return writerSheetBuilder.build();
    }

    private void fillCustomHeadInfo(Class<?> dataClass, Class<? extends HeadGenerator> headEnhancerClass,
                                    ExcelWriterSheetBuilder writerSheetBuilder) {
        HeadGenerator headGenerator = this.applicationContext.getBean(headEnhancerClass);
        Assert.notNull(headGenerator, "The header generated bean does not exist.");
        HeadMeta head = headGenerator.head(dataClass);
        writerSheetBuilder.head(head.getHead());
        writerSheetBuilder.excludeColumnFieldNames(head.getIgnoreHeadFields());
    }

    /**
     * 获取没有 @ExcelProperty 注解的字段名列表
     *
     * @param dataClass 数据类型
     * @return 没有 @ExcelProperty 注解的字段名列表
     */
    private List<String> getFieldsWithoutExcelProperty(Class<?> dataClass) {
        List<String> fieldsWithoutAnnotation = new java.util.ArrayList<>();
        java.lang.reflect.Field[] fields = dataClass.getDeclaredFields();

        for (java.lang.reflect.Field field : fields) {
            // 检查字段是否有 @ExcelProperty 注解
            if (!field.isAnnotationPresent(com.alibaba.excel.annotation.ExcelProperty.class)) {
                fieldsWithoutAnnotation.add(field.getName());
            }
        }

        return fieldsWithoutAnnotation;
    }

    /**
     * 是否为Null Head Generator
     *
     * @param headGeneratorClass 头生成器类型
     * @return true 已指定 false 未指定(默认值)
     */
    private boolean isNotInterface(Class<? extends HeadGenerator> headGeneratorClass) {
        return !Modifier.isInterface(headGeneratorClass.getModifiers());
    }

    /**
     * 检查实体类是否有 @ExcelValidation 注解
     *
     * @param dataClass 数据类型
     * @return true 有注解 false 无注解
     */
    private boolean hasExcelValidationAnnotation(Class<?> dataClass) {
        java.lang.reflect.Field[] fields = dataClass.getDeclaredFields();
        for (java.lang.reflect.Field field : fields) {
            if (field.isAnnotationPresent(cn.allbs.excel.annotation.ExcelValidation.class)) {
                return true;
            }
        }
        return false;
    }

    private boolean hasConditionalStyleAnnotation(Class<?> dataClass) {
        java.lang.reflect.Field[] fields = dataClass.getDeclaredFields();
        for (java.lang.reflect.Field field : fields) {
            if (field.isAnnotationPresent(cn.allbs.excel.annotation.ConditionalStyle.class)) {
                return true;
            }
        }
        return false;
    }

    private boolean hasExcelImageAnnotation(Class<?> dataClass) {
        java.lang.reflect.Field[] fields = dataClass.getDeclaredFields();
        for (java.lang.reflect.Field field : fields) {
            if (field.isAnnotationPresent(cn.allbs.excel.annotation.ExcelImage.class)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }


}
