package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.CrossFieldExpression;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.List;

import org.springframework.expression.Expression;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

/**
 * SpEL 表达式校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class CrossFieldExpressionRule implements CrossValidationRule {

    private static final ExpressionParser PARSER = new SpelExpressionParser();

    private final CrossFieldExpression annotation;
    private final Expression compiledExpression;

    public CrossFieldExpressionRule(CrossFieldExpression annotation) {
        this.annotation = annotation;
        // 预编译表达式
        this.compiledExpression = PARSER.parseExpression(annotation.expression());
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        try {
            StandardEvaluationContext context = new StandardEvaluationContext(target);

            // 将声明的字段注册为变量
            for (String fieldName : annotation.fields()) {
                Object value = FieldAccessorCache.getFieldValue(target, fieldName);
                context.setVariable(fieldName, value);
            }

            // 注册根对象
            context.setVariable("root", target);

            // 执行表达式
            Boolean result = compiledExpression.getValue(context, Boolean.class);

            if (result == null || !result) {
                errors.add(buildFieldError());
            }
        } catch (Exception e) {
            // 表达式执行失败，记录错误
            errors.add(FieldError.builder()
                    .fieldName("expression")
                    .errorType("CrossFieldExpression")
                    .message("表达式执行失败: " + e.getMessage())
                    .fullMessage("表达式 [" + annotation.expression() + "] 执行失败: " + e.getMessage())
                    .build());
        }

        return errors;
    }

    /**
     * 构建错误信息
     */
    private FieldError buildFieldError() {
        String[] fields = annotation.fields();
        String fieldNames = fields.length > 0 ? String.join(",", fields) : "expression";

        return FieldError.builder()
                .fieldName(fieldNames)
                .propertyName(fieldNames)
                .errorType("CrossFieldExpression")
                .message(annotation.message())
                .fullMessage(annotation.message())
                .build();
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}
