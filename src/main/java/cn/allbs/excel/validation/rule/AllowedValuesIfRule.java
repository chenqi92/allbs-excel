package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.AllowedValuesIf;
import cn.allbs.excel.annotation.cross.Condition;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

/**
 * 值依赖校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class AllowedValuesIfRule implements CrossValidationRule {

    private final String fieldName;
    private final AllowedValuesIf annotation;
    private final Pattern compiledPattern;

    public AllowedValuesIfRule(String fieldName, AllowedValuesIf annotation) {
        this.fieldName = fieldName;
        this.annotation = annotation;
        // 预编译正则表达式
        String patternStr = annotation.pattern();
        this.compiledPattern = (patternStr != null && !patternStr.isEmpty())
                ? Pattern.compile(patternStr)
                : null;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        // 检查条件是否满足
        boolean conditionMet = checkCondition(target);

        if (conditionMet) {
            // 条件满足时，检查当前字段值是否在允许范围内
            Object fieldValue = FieldAccessorCache.getFieldValue(target, fieldName);

            // 空值跳过校验
            if (ValidationHelper.isEmpty(fieldValue)) {
                return errors;
            }

            if (!isValueAllowed(fieldValue)) {
                errors.add(ValidationHelper.buildFieldError(
                        target.getClass(), fieldName, "AllowedValuesIf", annotation.message(), fieldValue));
            }
        }

        return errors;
    }

    /**
     * 检查条件是否满足
     */
    private boolean checkCondition(Object target) {
        Condition[] conditions = annotation.conditions();

        // 多条件模式
        if (conditions.length > 0) {
            return ValidationHelper.checkMultipleConditions(target, conditions);
        }

        // 单条件模式
        String dependField = annotation.dependField();
        if (dependField.isEmpty()) {
            return true; // 没有依赖条件，始终执行校验
        }

        String dependValue = annotation.dependValue();
        Object actualValue = FieldAccessorCache.getFieldValue(target, dependField);

        if (dependValue.isEmpty()) {
            return ValidationHelper.isNotEmpty(actualValue);
        } else {
            // 安全比较，处理 null 值
            return ValidationHelper.safeEquals(dependValue, actualValue);
        }
    }

    /**
     * 检查值是否在允许范围内
     */
    private boolean isValueAllowed(Object value) {
        String[] allowedValues = annotation.allowedValues();

        // 检查枚举值
        if (allowedValues.length > 0) {
            String strValue = String.valueOf(value);
            return Arrays.asList(allowedValues).contains(strValue);
        }

        // 检查正则
        if (compiledPattern != null) {
            String strValue = String.valueOf(value);
            return compiledPattern.matcher(strValue).matches();
        }

        // 检查数值范围
        if (value instanceof Number) {
            double numValue = ((Number) value).doubleValue();
            double minValue = annotation.minValue();
            double maxValue = annotation.maxValue();

            if (minValue != -Double.MAX_VALUE && numValue < minValue) {
                return false;
            }
            if (maxValue != Double.MAX_VALUE && numValue > maxValue) {
                return false;
            }
        }

        return true;
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}
