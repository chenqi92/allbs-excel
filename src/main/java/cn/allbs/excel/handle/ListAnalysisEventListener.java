package cn.allbs.excel.handle;

import cn.allbs.excel.vo.ErrorMessage;
import cn.idev.excel.event.AnalysisEventListener;

import java.util.List;

/**
 * @author ChenQi
 */
public abstract class ListAnalysisEventListener<T> extends AnalysisEventListener<T> {

    /**
     * 获取 excel 解析的对象列表
     *
     * @return 集合
     */
    public abstract List<T> getList();

    /**
     * 获取异常校验结果
     *
     * @return 集合
     */
    public abstract List<ErrorMessage> getErrors();

}
