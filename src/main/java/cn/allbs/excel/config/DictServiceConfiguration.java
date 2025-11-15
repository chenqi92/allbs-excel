package cn.allbs.excel.config;

import cn.allbs.excel.convert.DictConverter;
import cn.allbs.excel.service.DictService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;

import javax.annotation.PostConstruct;

/**
 * 字典服务配置类
 * <p>
 * 用于自动注入字典服务到字典转换器中
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Slf4j
@Configuration
public class DictServiceConfiguration {

    @Autowired(required = false)
    private DictService dictService;

    @PostConstruct
    public void init() {
        if (dictService != null) {
            DictConverter.setDictService(dictService);
            log.info("DictService 已注入到 DictConverter，字典转换功能已启用");
        } else {
            log.warn("未找到 DictService 实现，字典转换功能不可用。如需使用字典转换功能，请实现 DictService 接口并注册为 Spring Bean");
        }
    }
}

