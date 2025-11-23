package cn.allbs.excel.processor;

import cn.allbs.excel.constant.StringPool;
import org.springframework.context.expression.MethodBasedEvaluationContext;
import org.springframework.core.DefaultParameterNameDiscoverer;
import org.springframework.core.ParameterNameDiscoverer;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.UUID;

/**
 * SpEL 表达式处理器
 * <p>
 * 支持的功能：
 * <ul>
 *   <li>方法参数访问：#{#paramName}</li>
 *   <li>静态方法调用：#{T(java.time.LocalDate).now()}</li>
 *   <li>预定义变量：#{#now}, #{#today}, #{#timestamp}, #{#uuid}</li>
 *   <li>自定义函数：#{#formatDate(#date, 'yyyyMMdd')}, #{#sanitize(#filename)}</li>
 * </ul>
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public class NameSpelExpressionProcessor implements NameProcessor {

    /**
     * 参数发现器
     */
    private static final ParameterNameDiscoverer NAME_DISCOVERER = new DefaultParameterNameDiscoverer();

    /**
     * Express语法解析器
     */
    private static final ExpressionParser PARSER = new SpelExpressionParser();

    @Override
    public String doDetermineName(Object[] args, Method method, String key) {

        if (!key.contains(StringPool.HASH)) {
            return key;
        }

        // 创建基于方法的上下文
        MethodBasedEvaluationContext context = new MethodBasedEvaluationContext(null, method, args, NAME_DISCOVERER);

        // 添加预定义变量
        addPredefinedVariables(context);

        // 注册自定义函数
        registerCustomFunctions(context);

        final Object value = PARSER.parseExpression(key).getValue(context);
        return value == null ? null : value.toString();
    }

    /**
     * 添加预定义变量
     *
     * @param context SpEL 上下文
     */
    private void addPredefinedVariables(EvaluationContext context) {
        if (context instanceof StandardEvaluationContext) {
            StandardEvaluationContext stdContext = (StandardEvaluationContext) context;

            // 当前日期时间
            stdContext.setVariable("now", LocalDateTime.now());

            // 当前日期
            stdContext.setVariable("today", LocalDate.now());

            // 时间戳（毫秒）
            stdContext.setVariable("timestamp", System.currentTimeMillis());

            // UUID
            stdContext.setVariable("uuid", UUID.randomUUID().toString());
        }
    }

    /**
     * 注册自定义函数
     *
     * @param context SpEL 上下文
     */
    private void registerCustomFunctions(EvaluationContext context) {
        if (context instanceof StandardEvaluationContext) {
            StandardEvaluationContext stdContext = (StandardEvaluationContext) context;

            try {
                // 注册日期格式化函数
                stdContext.registerFunction("formatDate",
                        NameSpelExpressionProcessor.class.getDeclaredMethod("formatDate", LocalDate.class, String.class));

                // 注册日期时间格式化函数
                stdContext.registerFunction("formatDateTime",
                        NameSpelExpressionProcessor.class.getDeclaredMethod("formatDateTime", LocalDateTime.class, String.class));

                // 注册文件名清理函数
                stdContext.registerFunction("sanitize",
                        NameSpelExpressionProcessor.class.getDeclaredMethod("sanitize", String.class));


            } catch (NoSuchMethodException e) {
                // 忽略注册失败
            }
        }
    }

    /**
     * 格式化日期
     *
     * @param date    日期
     * @param pattern 格式模式
     * @return 格式化后的字符串
     */
    public static String formatDate(LocalDate date, String pattern) {
        if (date == null) {
            return "";
        }
        return date.format(DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * 格式化日期时间
     *
     * @param dateTime 日期时间
     * @param pattern  格式模式
     * @return 格式化后的字符串
     */
    public static String formatDateTime(LocalDateTime dateTime, String pattern) {
        if (dateTime == null) {
            return "";
        }
        return dateTime.format(DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * 清理文件名中的非法字符
     * <p>
     * 移除或替换 Windows 和 Unix 系统中不允许的文件名字符：\ / : * ? " < > |
     * </p>
     *
     * @param filename 原始文件名
     * @return 清理后的文件名
     */
    public static String sanitize(String filename) {
        if (filename == null) {
            return "";
        }
        return filename.replaceAll("[\\\\/:*?\"<>|]", "_");
    }
}
