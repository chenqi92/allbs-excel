package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.ConditionalPattern;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.FieldError;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * 条件正则校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class ConditionalPatternRule implements CrossValidationRule {

    private final String fieldName;
    private final ConditionalPattern annotation;
    private final Map<String, Pattern> compiledPatterns;
    private final Pattern defaultPattern;

    public ConditionalPatternRule(String fieldName, ConditionalPattern annotation) {
        this.fieldName = fieldName;
        this.annotation = annotation;

        // 预编译所有正则表达式
        this.compiledPatterns = new HashMap<>();
        for (ConditionalPattern.PatternRule rule : annotation.patterns()) {
            compiledPatterns.put(rule.value(), Pattern.compile(rule.pattern()));
        }

        String defaultPatternStr = annotation.defaultPattern();
        this.defaultPattern = (defaultPatternStr != null && !defaultPatternStr.isEmpty())
                ? Pattern.compile(defaultPatternStr) : null;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        Object fieldValue = FieldAccessorCache.getFieldValue(target, fieldName);

        // 空值跳过校验
        if (FieldAccessorCache.isEmpty(fieldValue)) {
            return errors;
        }

        Object dependValue = FieldAccessorCache.getFieldValue(target, annotation.dependField());
        String dependValueStr = dependValue != null ? String.valueOf(dependValue) : "";

        // 获取对应的正则
        Pattern pattern = compiledPatterns.get(dependValueStr);
        if (pattern == null) {
            pattern = defaultPattern;
        }

        // 如果没有匹配的正则，跳过校验
        if (pattern == null) {
            return errors;
        }

        // 执行正则校验
        String fieldValueStr = String.valueOf(fieldValue);
        if (!pattern.matcher(fieldValueStr).matches()) {
            errors.add(buildFieldError(target, fieldValue));
        }

        return errors;
    }

    /**
     * 构建错误信息
     */
    private FieldError buildFieldError(Object target, Object actualValue) {
        String excelFieldName = getExcelFieldName(target.getClass(), fieldName);
        return FieldError.builder()
                .fieldName(excelFieldName)
                .propertyName(fieldName)
                .errorType("ConditionalPattern")
                .message(annotation.message())
                .fullMessage("【" + excelFieldName + "】" + annotation.message())
                .fieldValue(actualValue)
                .build();
    }

    /**
     * 获取 Excel 列名
     */
    private String getExcelFieldName(Class<?> clazz, String fieldName) {
        return ValidationHelper.getExcelFieldName(clazz, fieldName);
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}
