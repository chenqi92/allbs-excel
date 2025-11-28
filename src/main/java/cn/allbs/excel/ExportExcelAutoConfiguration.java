package cn.allbs.excel;

import cn.allbs.excel.aop.DynamicNameAspect;
import cn.allbs.excel.aop.ExportExcelReturnValueHandler;
import cn.allbs.excel.aop.ImportExcelArgumentResolver;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.processor.NameProcessor;
import cn.allbs.excel.processor.NameSpelExpressionProcessor;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.beans.factory.SmartInitializingSingleton;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Import;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;

/**
 * 配置初始化
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29 16:03
 */
@AutoConfiguration
@Import({ExcelHandlerConfiguration.class, cn.allbs.excel.config.DictServiceConfiguration.class})
@EnableConfigurationProperties(ExcelConfigProperties.class)
public class ExportExcelAutoConfiguration implements SmartInitializingSingleton {

    private final RequestMappingHandlerAdapter requestMappingHandlerAdapter;

    private final ExportExcelReturnValueHandler exportExcelReturnValueHandler;

    private final ObjectProvider<ImportExcelArgumentResolver> importExcelArgumentResolverProvider;

    public ExportExcelAutoConfiguration(RequestMappingHandlerAdapter requestMappingHandlerAdapter,
                                        ExportExcelReturnValueHandler exportExcelReturnValueHandler,
                                        ObjectProvider<ImportExcelArgumentResolver> importExcelArgumentResolverProvider) {
        this.requestMappingHandlerAdapter = requestMappingHandlerAdapter;
        this.exportExcelReturnValueHandler = exportExcelReturnValueHandler;
        this.importExcelArgumentResolverProvider = importExcelArgumentResolverProvider;
    }

    /**
     * SPEL 解析处理器
     *
     * @return NameProcessor excel名称解析器
     */
    @Bean
    @ConditionalOnMissingBean
    public NameProcessor nameProcessor() {
        return new NameSpelExpressionProcessor();
    }

    /**
     * Excel名称解析处理切面
     *
     * @param nameProcessor SPEL 解析处理器
     * @return DynamicNameAspect
     */
    @Bean
    @ConditionalOnMissingBean
    public DynamicNameAspect dynamicNameAspect(NameProcessor nameProcessor) {
        return new DynamicNameAspect(nameProcessor);
    }

    /**
     * 追加 Excel返回值处理器 到 springmvc 中
     */
    @PostConstruct
    public void setReturnValueHandlers() {
        List<HandlerMethodReturnValueHandler> returnValueHandlers = requestMappingHandlerAdapter
                .getReturnValueHandlers();

        List<HandlerMethodReturnValueHandler> newHandlers = new ArrayList<>();
        newHandlers.add(exportExcelReturnValueHandler);
        assert returnValueHandlers != null;
        newHandlers.addAll(returnValueHandlers);
        requestMappingHandlerAdapter.setReturnValueHandlers(newHandlers);
    }

    /**
     * 创建 ImportExcelArgumentResolver Bean
     * 必须注册为Bean才能使ApplicationContextAware生效
     */
    @Bean
    @ConditionalOnMissingBean
    public ImportExcelArgumentResolver importExcelArgumentResolver() {
        return new ImportExcelArgumentResolver();
    }

    /**
     * 追加 Excel 请求处理器 到 springmvc 中
     * 使用 SmartInitializingSingleton 确保所有单例bean都已完全初始化
     */
    @Override
    public void afterSingletonsInstantiated() {
        List<HandlerMethodArgumentResolver> argumentResolvers = requestMappingHandlerAdapter.getArgumentResolvers();
        List<HandlerMethodArgumentResolver> resolverList = new ArrayList<>();
        // 使用 ObjectProvider 获取Spring管理的bean实例，确保ApplicationContextAware生效
        ImportExcelArgumentResolver resolver = importExcelArgumentResolverProvider.getIfAvailable();
        if (resolver != null) {
            resolverList.add(resolver);
        } else {
            // 如果bean不可用，使用新实例（兼容旧行为）
            resolverList.add(new ImportExcelArgumentResolver());
        }
        assert argumentResolvers != null;
        resolverList.addAll(argumentResolvers);
        requestMappingHandlerAdapter.setArgumentResolvers(resolverList);
    }

}
