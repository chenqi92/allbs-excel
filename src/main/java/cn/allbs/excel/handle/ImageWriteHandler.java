package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelImage;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URL;
import java.util.*;

/**
 * 图片写入处理器
 * <p>
 * 支持将字段值作为图片导出到 Excel，支持多种图片来源：URL、本地文件、字节数组、Base64
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ImageWriteHandler implements CellWriteHandler, WorkbookWriteHandler {

	/**
	 * 数据类型
	 */
	private Class<?> dataClass;

	/**
	 * 列索引 -> 图片配置信息
	 */
	private final Map<Integer, ImageColumnInfo> imageColumnMap = new HashMap<>();

	/**
	 * 待插入的图片信息（Sheet名 -> 图片列表）
	 */
	private final Map<String, List<ImageInfo>> pendingImages = new HashMap<>();

	/**
	 * 表头行数
	 */
	private int headRowNumber = 1;

	/**
	 * Default constructor (for reflection instantiation)
	 */
	public ImageWriteHandler() {
		this.dataClass = null;
	}

	public ImageWriteHandler(Class<?> dataClass) {
		this.dataClass = dataClass;
		initImageColumns();
	}

	/**
	 * 初始化图片列配置
	 */
	private void initImageColumns() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			ExcelImage excelImage = field.getAnnotation(ExcelImage.class);

			if (excelProperty != null) {
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				if (excelImage != null) {
					ImageColumnInfo info = new ImageColumnInfo(field, excelImage);
					imageColumnMap.put(index, info);
					log.debug("Registered image column {}: {}", index, field.getName());
				}

				columnIndex++;
			}
		}

		log.info("Initialized {} image columns", imageColumnMap.size());
	}

	@Override
	public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
								  List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
		// 记录表头行数
		if (isHead != null && isHead) {
			headRowNumber = Math.max(headRowNumber, cell.getRowIndex() + 1);
			return;
		}

		int columnIndex = cell.getColumnIndex();
		ImageColumnInfo imageInfo = imageColumnMap.get(columnIndex);

		if (imageInfo == null) {
			return;
		}

		// 从 cellDataList 中获取原始数据
		if (cellDataList == null || cellDataList.isEmpty()) {
			return;
		}

		Object cellValue = cellDataList.get(0).getData();
		if (cellValue == null) {
			return;
		}

		// 处理图片数据
		String sheetName = writeSheetHolder.getSheet().getSheetName();
		int rowIndex = cell.getRowIndex();

		try {
			List<byte[]> imageDataList = loadImageData(cellValue, imageInfo);

			if (!imageDataList.isEmpty()) {
				for (int i = 0; i < imageDataList.size(); i++) {
					byte[] imageData = imageDataList.get(i);
					ImageInfo info = new ImageInfo(sheetName, rowIndex, columnIndex, imageData,
							imageInfo.excelImage.width(), imageInfo.excelImage.height(), i, imageDataList.size());
					pendingImages.computeIfAbsent(sheetName, k -> new ArrayList<>()).add(info);
				}

				// 清空单元格内容（因为图片会覆盖）
				cell.setCellValue("");
				log.debug("Scheduled {} image(s) for cell [{}, {}]", imageDataList.size(), rowIndex, columnIndex);
			}
		}
		catch (Exception e) {
			log.error("Failed to load image data for cell [{}, {}]: {}", rowIndex, columnIndex, e.getMessage());
			if (imageInfo.excelImage.showPlaceholderOnError()) {
				cell.setCellValue("图片加载失败");
			}
		}
	}

	/**
	 * 加载图片数据
	 *
	 * @param cellValue  单元格值
	 * @param columnInfo 列信息
	 * @return 图片数据列表
	 */
	private List<byte[]> loadImageData(Object cellValue, ImageColumnInfo columnInfo) throws Exception {
		List<byte[]> result = new ArrayList<>();

		// 如果是集合类型，遍历处理
		if (cellValue instanceof Collection) {
			Collection<?> collection = (Collection<?>) cellValue;
			for (Object item : collection) {
				byte[] data = loadSingleImage(item, columnInfo.excelImage);
				if (data != null) {
					result.add(data);
				}
			}
		}
		else {
			byte[] data = loadSingleImage(cellValue, columnInfo.excelImage);
			if (data != null) {
				result.add(data);
			}
		}

		return result;
	}

	/**
	 * 加载单张图片
	 *
	 * @param value      图片值
	 * @param excelImage 图片配置
	 * @return 图片字节数组
	 */
	private byte[] loadSingleImage(Object value, ExcelImage excelImage) throws Exception {
		if (value == null) {
			return null;
		}

		// 如果已经是字节数组，直接返回
		if (value instanceof byte[]) {
			return (byte[]) value;
		}

		String valueStr = value.toString().trim();
		if (valueStr.isEmpty()) {
			return null;
		}

		ExcelImage.ImageType type = excelImage.type();

		// 自动检测类型
		if (type == ExcelImage.ImageType.AUTO) {
			type = detectImageType(valueStr);
		}

		switch (type) {
			case URL:
				return loadImageFromUrl(valueStr);
			case FILE:
				return loadImageFromFile(valueStr);
			case BASE64:
				return loadImageFromBase64(valueStr);
			case BYTES:
				return (byte[]) value;
			default:
				log.warn("Unknown image type: {}", type);
				return null;
		}
	}

	/**
	 * 自动检测图片类型
	 */
	private ExcelImage.ImageType detectImageType(String value) {
		if (value.startsWith("http://") || value.startsWith("https://")) {
			return ExcelImage.ImageType.URL;
		}
		else if (value.startsWith("data:image/")) {
			return ExcelImage.ImageType.BASE64;
		}
		else if (new File(value).exists()) {
			return ExcelImage.ImageType.FILE;
		}
		else {
			return ExcelImage.ImageType.FILE; // 默认作为文件路径
		}
	}

	/**
	 * 从URL加载图片
	 */
	private byte[] loadImageFromUrl(String url) throws Exception {
		try (InputStream is = new URL(url).openStream(); ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
			byte[] buffer = new byte[4096];
			int bytesRead;
			while ((bytesRead = is.read(buffer)) != -1) {
				baos.write(buffer, 0, bytesRead);
			}
			return baos.toByteArray();
		}
	}

	/**
	 * 从本地文件加载图片
	 */
	private byte[] loadImageFromFile(String filePath) throws Exception {
		File file = new File(filePath);
		if (!file.exists()) {
			throw new FileNotFoundException("Image file not found: " + filePath);
		}

		try (FileInputStream fis = new FileInputStream(file); ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
			byte[] buffer = new byte[4096];
			int bytesRead;
			while ((bytesRead = fis.read(buffer)) != -1) {
				baos.write(buffer, 0, bytesRead);
			}
			return baos.toByteArray();
		}
	}

	/**
	 * 从Base64字符串加载图片
	 */
	private byte[] loadImageFromBase64(String base64Str) {
		// 移除 data:image/xxx;base64, 前缀
		if (base64Str.contains(",")) {
			base64Str = base64Str.substring(base64Str.indexOf(",") + 1);
		}
		return Base64.getDecoder().decode(base64Str);
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
		// 在工作簿完全写入后，统一插入所有图片
		if (pendingImages.isEmpty()) {
			log.debug("No pending images to insert");
			return;
		}

		log.info("Inserting {} images into workbook", pendingImages.values().stream().mapToInt(List::size).sum());

		Workbook workbook = writeWorkbookHolder.getWorkbook();
		Drawing<?> drawing = null;
		CreationHelper helper = workbook.getCreationHelper();

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			List<ImageInfo> sheetImages = pendingImages.get(sheetName);

			if (sheetImages == null || sheetImages.isEmpty()) {
				continue;
			}

			// 创建绘图对象
			drawing = sheet.createDrawingPatriarch();

			log.debug("Inserting {} images into sheet: {}", sheetImages.size(), sheetName);

			for (ImageInfo imageInfo : sheetImages) {
				try {
					insertImage(workbook, sheet, drawing, helper, imageInfo);
				}
				catch (Exception e) {
					log.error("Failed to insert image at [{}, {}]: {}", imageInfo.rowIndex, imageInfo.columnIndex,
							e.getMessage());
				}
			}
		}

		log.info("Finished inserting images");
	}

	/**
	 * 插入图片到单元格
	 */
	private void insertImage(Workbook workbook, Sheet sheet, Drawing<?> drawing, CreationHelper helper,
			ImageInfo imageInfo) throws Exception {

		// 添加图片到工作簿
		int pictureIdx = workbook.addPicture(imageInfo.imageData, Workbook.PICTURE_TYPE_PNG);

		// 计算图片位置
		ClientAnchor anchor = helper.createClientAnchor();

		// 设置图片位置：从指定单元格开始
		anchor.setCol1(imageInfo.columnIndex);
		anchor.setRow1(imageInfo.rowIndex);

		// 如果一个单元格有多张图片，需要调整位置
		int offsetX = 0;
		int offsetY = 0;

		if (imageInfo.totalImages > 1) {
			// 水平排列多张图片
			offsetX = imageInfo.imageIndex * (imageInfo.width + 5); // 5像素间距
		}

		// 设置图片的起始位置偏移（以EMU为单位，1像素 ≈ 9525 EMU）
		anchor.setDx1(offsetX * Units.EMU_PER_PIXEL);
		anchor.setDy1(offsetY * Units.EMU_PER_PIXEL);

		// 设置图片的结束位置
		anchor.setCol2(imageInfo.columnIndex + 1);
		anchor.setRow2(imageInfo.rowIndex + 1);
		anchor.setDx2((offsetX + imageInfo.width) * Units.EMU_PER_PIXEL);
		anchor.setDy2(imageInfo.height * Units.EMU_PER_PIXEL);

		// 设置 anchor 类型为 DONT_MOVE_AND_RESIZE，确保图片不会随单元格大小变化而变形
		// 这可以防止图片被压缩或拉伸
		anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

		// 创建图片
		Picture picture = drawing.createPicture(anchor, pictureIdx);

		// 调整行高以适应图片（确保行高足够显示图片）
		Row row = sheet.getRow(imageInfo.rowIndex);
		if (row != null) {
			// 将像素转换为行高单位（1行高单位 = 1/20磅，1磅 ≈ 1.33像素）
			// 添加一些额外空间（+10）以避免图片被压缩
			short requiredHeight = (short) (((imageInfo.height + 10) * 20) / 1.33);
			if (row.getHeight() < requiredHeight) {
				row.setHeight(requiredHeight);
			}
		}

		// 调整列宽以适应图片
		int totalWidth = imageInfo.totalImages > 1
				? (imageInfo.width + 5) * imageInfo.totalImages
				: imageInfo.width;

		// 列宽单位：1个单位 = 1/256个字符宽度
		// 添加额外空间（+15）以确保图片完全显示
		int currentWidth = sheet.getColumnWidth(imageInfo.columnIndex);
		int newWidth = ((totalWidth + 15) * 256) / 7; // 假设字符宽度约为7像素
		if (currentWidth < newWidth) {
			sheet.setColumnWidth(imageInfo.columnIndex, newWidth);
		}

		log.trace("Inserted image at cell [{}, {}] with size {}x{}", imageInfo.rowIndex, imageInfo.columnIndex,
				imageInfo.width, imageInfo.height);
	}

	/**
	 * 图片列信息
	 */
	private static class ImageColumnInfo {

		private final Field field;

		private final ExcelImage excelImage;

		public ImageColumnInfo(Field field, ExcelImage excelImage) {
			this.field = field;
			this.excelImage = excelImage;
		}

	}

	/**
	 * 图片信息
	 */
	private static class ImageInfo {

		private final String sheetName;

		private final int rowIndex;

		private final int columnIndex;

		private final byte[] imageData;

		private final int width;

		private final int height;

		private final int imageIndex; // 在同一单元格中的索引（从0开始）

		private final int totalImages; // 同一单元格中的总图片数

		public ImageInfo(String sheetName, int rowIndex, int columnIndex, byte[] imageData, int width, int height,
				int imageIndex, int totalImages) {
			this.sheetName = sheetName;
			this.rowIndex = rowIndex;
			this.columnIndex = columnIndex;
			this.imageData = imageData;
			this.width = width;
			this.height = height;
			this.imageIndex = imageIndex;
			this.totalImages = totalImages;
		}

	}

}
