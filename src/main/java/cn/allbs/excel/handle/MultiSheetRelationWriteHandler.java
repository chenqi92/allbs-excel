package cn.allbs.excel.handle;

import cn.allbs.excel.util.MultiSheetRelationProcessor;
import cn.idev.excel.write.handler.WorkbookWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * 多 Sheet 关联写入处理器
 * <p>
 * 用于在 EasyExcel 完成写入后应用超链接
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class MultiSheetRelationWriteHandler implements WorkbookWriteHandler {

	/**
	 * 待应用的超链接列表
	 */
	private final List<MultiSheetRelationProcessor.HyperlinkInfo> hyperlinks;

	public MultiSheetRelationWriteHandler(List<MultiSheetRelationProcessor.HyperlinkInfo> hyperlinks) {
		this.hyperlinks = hyperlinks;
	}

	@Override
	public void beforeWorkbookCreate() {
		// 无需处理
	}

	@Override
	public void afterWorkbookCreate(WriteWorkbookHolder writeWorkbookHolder) {
		// 无需处理
	}

	@Override
	public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
		// 在工作簿完全写入后，应用所有超链接
		if (hyperlinks == null || hyperlinks.isEmpty()) {
			log.debug("No hyperlinks to apply");
			return;
		}

		Workbook workbook = writeWorkbookHolder.getWorkbook();
		MultiSheetRelationProcessor.applyHyperlinks(workbook, hyperlinks);
	}

}
