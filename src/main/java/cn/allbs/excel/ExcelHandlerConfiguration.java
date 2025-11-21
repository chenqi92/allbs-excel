package cn.allbs.excel;

import cn.allbs.excel.aop.ExportExcelReturnValueHandler;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.enhance.DefaultWriterBuilderEnhancer;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.handle.*;
import cn.allbs.excel.head.I18nHeaderCellWriteHandler;
import com.alibaba.excel.converters.Converter;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.boot.autoconfigure.condition.ConditionalOnBean;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.MessageSource;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.annotation.Order;

import java.util.List;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29 16:03
 */
@Configuration
@RequiredArgsConstructor
public class ExcelHandlerConfiguration {
    private final ExcelConfigProperties configProperties;

    private final ObjectProvider<List<Converter<?>>> converterProvider;

    /**
     * ExcelBuild增强
     *
     * @return DefaultWriterBuilderEnhancer 默认什么也不做的增强器
     */
    @Bean
    @ConditionalOnMissingBean
    public WriterBuilderEnhancer writerBuilderEnhancer() {
        return new DefaultWriterBuilderEnhancer();
    }

    /**
     * 动态表头写入处理器（优先级最高）
     * 用于处理带有 @DynamicHeaders 注解的实体类
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(1)
    public DynamicHeaderWriteHandler dynamicHeaderWriteHandler() {
        return new DynamicHeaderWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * FlattenList 写入处理器
     * 用于处理带有 @FlattenList 注解的实体类
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(2)
    public FlattenListWriteHandler flattenListWriteHandler() {
        return new FlattenListWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * FlattenProperty 写入处理器
     * 用于处理带有 @FlattenProperty 注解的实体类
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(3)
    public FlattenPropertyWriteHandler flattenPropertyWriteHandler() {
        return new FlattenPropertyWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * 关联 Sheet 写入处理器
     * 用于处理带有 @RelatedSheet 注解的实体类
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(4)
    public RelatedSheetWriteHandler relatedSheetWriteHandler() {
        return new RelatedSheetWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * 多sheet 写入处理器
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(5)
    public ManySheetWriteHandler manySheetWriteHandler() {
        return new ManySheetWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * 单sheet 写入处理器（优先级最低，兜底）
     */
    @Bean
    @ConditionalOnMissingBean
    @Order(6)
    public SingleSheetWriteHandler singleSheetWriteHandler() {
        return new SingleSheetWriteHandler(configProperties, converterProvider, writerBuilderEnhancer());
    }

    /**
     * 返回Excel文件的 response 处理器
     *
     * @param sheetWriteHandlerList 页签写入处理器集合
     * @return ResponseExcelReturnValueHandler
     */
    @Bean
    @ConditionalOnMissingBean
    public ExportExcelReturnValueHandler responseExcelReturnValueHandler(List<SheetWriteHandler> sheetWriteHandlerList) {
        return new ExportExcelReturnValueHandler(sheetWriteHandlerList);
    }

    /**
     * excel 头的国际化处理器
     *
     * @param messageSource 国际化源
     */
    @Bean
    @ConditionalOnBean(MessageSource.class)
    @ConditionalOnMissingBean
    public I18nHeaderCellWriteHandler i18nHeaderCellWriteHandler(MessageSource messageSource) {
        return new I18nHeaderCellWriteHandler(messageSource);
    }

}
