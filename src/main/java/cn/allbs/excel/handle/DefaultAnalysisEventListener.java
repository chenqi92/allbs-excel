package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelLine;
import cn.allbs.excel.kit.Validators;
import cn.allbs.excel.vo.ErrorMessage;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 默认的 AnalysisEventListener
 *
 * @author ChenQi
 */
@Slf4j
public class DefaultAnalysisEventListener extends ListAnalysisEventListener<Object> {

    private final List<Object> list = new ArrayList<>();

    private final List<ErrorMessage> errorMessageList = new ArrayList<>();

    private Long lineNum = 1L;

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        lineNum++;

        Set<ConstraintViolation<Object>> violations = Validators.validate(o);
        if (!violations.isEmpty()) {

            Set<String> errorSet = new HashSet<>();
            violations.forEach(a -> {
                try {
                    Field field = o.getClass().getDeclaredField(a.getPropertyPath().toString());
                    String filedName = Optional.ofNullable(field.getAnnotation(ExcelProperty.class)).map(anno -> anno.value()[0]).orElse(field.getName());
                    errorSet.add(filedName + a.getMessage());
                } catch (Exception e) {
                    log.error("域解析错误" + e.getLocalizedMessage());
                }
            });
            errorMessageList.add(new ErrorMessage(lineNum, errorSet));
        } else {
            Field[] fields = o.getClass().getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelLine.class) && field.getType() == Long.class) {
                    try {
                        field.setAccessible(true);
                        field.set(o, lineNum);
                    } catch (IllegalAccessException e) {
                        log.error("设置行号失败: {}", e.getMessage(), e);
                    }
                }
            }
            list.add(o);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed");
    }

    @Override
    public List<Object> getList() {
        return list;
    }

    @Override
    public List<ErrorMessage> getErrors() {
        return errorMessageList;
    }

}
