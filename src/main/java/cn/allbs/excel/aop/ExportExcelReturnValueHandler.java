package cn.allbs.excel.aop;

import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.handle.SheetWriteHandler;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.MethodParameter;
import org.springframework.util.Assert;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

import jakarta.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
@Slf4j
@RequiredArgsConstructor
public class ExportExcelReturnValueHandler implements HandlerMethodReturnValueHandler {
    private final List<SheetWriteHandler> sheetWriteHandlerList;


    /**
     * 只处理@ExportExcel 声明的方法
     *
     * @param parameter 方法签名
     * @return 是否处理
     */
    @Override
    public boolean supportsReturnType(MethodParameter parameter) {
        return parameter.getMethodAnnotation(ExportExcel.class) != null;
    }

    /**
     * 处理逻辑
     *
     * @param o                返回参数
     * @param parameter        方法签名
     * @param mavContainer     上下文容器
     * @param nativeWebRequest 上下文
     * @throws Exception 处理异常
     */
    @Override
    public void handleReturnValue(Object o, MethodParameter parameter, ModelAndViewContainer mavContainer,
                                  NativeWebRequest nativeWebRequest) {
        /* check */
        HttpServletResponse response = nativeWebRequest.getNativeResponse(HttpServletResponse.class);
        Assert.state(response != null, "No HttpServletResponse");
        ExportExcel responseExcel = parameter.getMethodAnnotation(ExportExcel.class);
        Assert.state(responseExcel != null, "No @ResponseExcel");
        mavContainer.setRequestHandled(true);

        sheetWriteHandlerList.stream().filter(handler -> handler.support(o)).findFirst()
                .ifPresent(handler -> handler.export(o, response, responseExcel));
    }
}
