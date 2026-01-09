package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.CompareOperator;
import cn.allbs.excel.annotation.cross.FieldsCompare;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.FieldError;
import cn.idev.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.temporal.Temporal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * 字段比较校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class FieldsCompareRule implements CrossValidationRule {

    private final FieldsCompare annotation;

    public FieldsCompareRule(FieldsCompare annotation) {
        this.annotation = annotation;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        Object firstValue = FieldAccessorCache.getFieldValue(target, annotation.first());
        Object secondValue = FieldAccessorCache.getFieldValue(target, annotation.second());

        // 处理空值
        if (firstValue == null && secondValue == null) {
            if (annotation.allowBothNull()) {
                return errors;
            }
        }

        if (firstValue == null || secondValue == null) {
            // 如果只有一个为空，无法比较
            return errors;
        }

        // 执行比较
        int compareResult = compare(firstValue, secondValue);
        boolean valid = checkOperator(compareResult, annotation.operator());

        if (!valid) {
            errors.add(buildFieldError(target));
        }

        return errors;
    }

    /**
     * 比较两个值
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private int compare(Object first, Object second) {
        // 数值比较
        if (first instanceof Number && second instanceof Number) {
            BigDecimal bd1 = new BigDecimal(first.toString());
            BigDecimal bd2 = new BigDecimal(second.toString());
            return bd1.compareTo(bd2);
        }

        // 日期比较
        if (first instanceof Comparable && second instanceof Comparable) {
            return ((Comparable) first).compareTo(second);
        }

        // 字符串比较
        return String.valueOf(first).compareTo(String.valueOf(second));
    }

    /**
     * 检查比较结果是否满足操作符
     */
    private boolean checkOperator(int compareResult, CompareOperator operator) {
        switch (operator) {
            case EQUAL:
                return compareResult == 0;
            case NOT_EQUAL:
                return compareResult != 0;
            case LESS_THAN:
                return compareResult < 0;
            case LESS_THAN_OR_EQUAL:
                return compareResult <= 0;
            case GREATER_THAN:
                return compareResult > 0;
            case GREATER_THAN_OR_EQUAL:
                return compareResult >= 0;
            default:
                return false;
        }
    }

    /**
     * 构建错误信息
     */
    private FieldError buildFieldError(Object target) {
        String firstExcelName = getExcelFieldName(target.getClass(), annotation.first());
        String secondExcelName = getExcelFieldName(target.getClass(), annotation.second());

        return FieldError.builder()
                .fieldName(firstExcelName + "," + secondExcelName)
                .propertyName(annotation.first() + "," + annotation.second())
                .errorType("FieldsCompare")
                .message(annotation.message())
                .fullMessage("【" + firstExcelName + "】" + annotation.operator().getDescription()
                        + "【" + secondExcelName + "】" + annotation.message())
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
