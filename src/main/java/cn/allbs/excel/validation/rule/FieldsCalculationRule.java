package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.CalculationOperator;
import cn.allbs.excel.annotation.cross.FieldsCalculation;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.FieldError;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 数值计算校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class FieldsCalculationRule implements CrossValidationRule {

    private final FieldsCalculation annotation;

    public FieldsCalculationRule(FieldsCalculation annotation) {
        this.annotation = annotation;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        String[] fields = annotation.fields();
        String resultField = annotation.resultField();

        // 获取所有字段值
        List<BigDecimal> values = new ArrayList<>();
        for (String fieldName : fields) {
            Object value = FieldAccessorCache.getFieldValue(target, fieldName);
            BigDecimal bdValue = toDecimal(value, annotation.ignoreNull());
            if (bdValue != null) {
                values.add(bdValue);
            } else if (!annotation.ignoreNull()) {
                // 如果不忽略空值且存在空值，跳过校验
                return errors;
            }
        }

        // 获取结果字段值
        Object resultValue = FieldAccessorCache.getFieldValue(target, resultField);
        if (resultValue == null) {
            return errors; // 结果字段为空，跳过校验
        }
        BigDecimal expectedResult = toDecimal(resultValue, false);
        if (expectedResult == null) {
            return errors;
        }

        // 计算
        BigDecimal calculatedResult = calculate(values, annotation.operator());
        if (calculatedResult == null) {
            return errors;
        }

        // 比较（考虑误差）
        double tolerance = annotation.tolerance();
        if (calculatedResult.subtract(expectedResult).abs().doubleValue() > tolerance) {
            errors.add(buildFieldError(target, calculatedResult, expectedResult));
        }

        return errors;
    }

    /**
     * 转换为 BigDecimal
     */
    private BigDecimal toDecimal(Object value, boolean ignoreNull) {
        if (value == null) {
            return ignoreNull ? BigDecimal.ZERO : null;
        }
        if (value instanceof BigDecimal) {
            return (BigDecimal) value;
        }
        if (value instanceof Number) {
            return new BigDecimal(value.toString());
        }
        try {
            return new BigDecimal(String.valueOf(value));
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /**
     * 执行计算
     */
    private BigDecimal calculate(List<BigDecimal> values, CalculationOperator operator) {
        if (values.isEmpty()) {
            return null;
        }

        switch (operator) {
            case SUM:
                return values.stream().reduce(BigDecimal.ZERO, BigDecimal::add);
            case MULTIPLY:
                return values.stream().reduce(BigDecimal.ONE, BigDecimal::multiply);
            case AVERAGE: {
                BigDecimal sum = values.stream().reduce(BigDecimal.ZERO, BigDecimal::add);
                return sum.divide(BigDecimal.valueOf(values.size()), 10, RoundingMode.HALF_UP);
            }
            case MAX:
                return values.stream().max(BigDecimal::compareTo).orElse(null);
            case MIN:
                return values.stream().min(BigDecimal::compareTo).orElse(null);
            case SUBTRACT: {
                BigDecimal first = values.get(0);
                for (int i = 1; i < values.size(); i++) {
                    first = first.subtract(values.get(i));
                }
                return first;
            }
            default:
                return null;
        }
    }

    /**
     * 构建错误信息
     */
    private FieldError buildFieldError(Object target, BigDecimal calculated, BigDecimal expected) {
        String resultExcelName = getExcelFieldName(target.getClass(), annotation.resultField());
        String fieldNames = Arrays.stream(annotation.fields())
                .map(f -> getExcelFieldName(target.getClass(), f))
                .collect(Collectors.joining("、"));

        return FieldError.builder()
                .fieldName(resultExcelName)
                .propertyName(annotation.resultField())
                .errorType("FieldsCalculation")
                .message(annotation.message())
                .fullMessage("【" + fieldNames + "】" + annotation.operator().getDescription()
                        + " 应为 " + calculated + "，实际为 " + expected)
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
