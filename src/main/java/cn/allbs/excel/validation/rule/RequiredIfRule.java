package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.Condition;
import cn.allbs.excel.annotation.cross.RequiredIf;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.List;

/**
 * 条件必填校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class RequiredIfRule implements CrossValidationRule {

    private final String fieldName;
    private final RequiredIf annotation;

    public RequiredIfRule(String fieldName, RequiredIf annotation) {
        this.fieldName = fieldName;
        this.annotation = annotation;
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
            // 条件满足时，检查当前字段是否有值
            Object fieldValue = FieldAccessorCache.getFieldValue(target, fieldName);
            if (ValidationHelper.isEmpty(fieldValue)) {
                errors.add(ValidationHelper.buildFieldError(
                        target.getClass(), fieldName, "RequiredIf", annotation.message()));
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
        String dependField = annotation.field();
        String hasValue = annotation.hasValue();

        Object dependValue = FieldAccessorCache.getFieldValue(target, dependField);

        if (hasValue.isEmpty()) {
            // 只要依赖字段非空即触发
            return ValidationHelper.isNotEmpty(dependValue);
        } else {
            // 依赖字段值必须等于指定值（安全比较，处理 null）
            return ValidationHelper.safeEquals(hasValue, dependValue);
        }
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}

