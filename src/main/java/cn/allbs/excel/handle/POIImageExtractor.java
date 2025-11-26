package cn.allbs.excel.handle;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * POI图片提取器
 * <p>
 * 直接使用POI API提取Excel中的Drawing对象（真实图片）
 * 绕过EasyExcel的限制，实现完整的图片读取功能
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-25
 */
@Slf4j
public class POIImageExtractor {

    /**
     * 提取的图片数据
     * 外层Map: sheetIndex -> 内层Map
     * 内层Map: "row_col" -> 图片数据列表
     */
    @Getter
    private final Map<Integer, Map<String, List<ImageData>>> extractedImages = new HashMap<>();

    /**
     * 从Excel文件中提取所有图片
     *
     * @param file Excel文件
     * @return 提取结果
     */
    public boolean extractImages(File file) {
        try (FileInputStream fis = new FileInputStream(file)) {
            return extractImages(fis, file.getName());
        } catch (Exception e) {
            log.error("提取图片失败: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * 从输入流中提取所有图片
     *
     * @param inputStream 输入流
     * @param fileName    文件名（用于判断格式）
     * @return 提取结果
     */
    public boolean extractImages(InputStream inputStream, String fileName) {
        try {
            Workbook workbook = WorkbookFactory.create(inputStream);

            if (workbook instanceof XSSFWorkbook) {
                extractXSSFImages((XSSFWorkbook) workbook);
            } else if (workbook instanceof HSSFWorkbook) {
                extractHSSFImages((HSSFWorkbook) workbook);
            } else {
                log.warn("不支持的Workbook类型: {}", workbook.getClass().getName());
                return false;
            }

            workbook.close();
            return true;
        } catch (Exception e) {
            log.error("提取图片失败: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * 提取XLSX格式的图片
     */
    private void extractXSSFImages(XSSFWorkbook workbook) {
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            Map<String, List<ImageData>> sheetImages = new HashMap<>();

            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            if (drawing == null) {
                log.debug("Sheet {} 没有Drawing对象", sheetIndex);
                continue;
            }

            List<XSSFShape> shapes = drawing.getShapes();
            log.info("Sheet {} 找到 {} 个图形对象", sheetIndex, shapes.size());

            for (XSSFShape shape : shapes) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getClientAnchor();

                    if (anchor != null) {
                        int row = anchor.getRow1();
                        int col = anchor.getCol1();
                        int row2 = anchor.getRow2();
                        int col2 = anchor.getCol2();
                        String key = row + "_" + col;

                        ImageData imageData = new ImageData();
                        XSSFPictureData pictureData = picture.getPictureData();
                        imageData.data = pictureData.getData();
                        imageData.mimeType = pictureData.getMimeType();
                        imageData.pictureType = pictureData.getPictureType();
                        imageData.row = row;
                        imageData.col = col;
                        imageData.row2 = row2;
                        imageData.col2 = col2;
                        imageData.fileName = pictureData.getPackagePart() != null
                            ? pictureData.getPackagePart().getPartName().getName()
                            : "image_" + row + "_" + col;

                        sheetImages.computeIfAbsent(key, k -> new ArrayList<>()).add(imageData);

                        log.debug("提取图片: Sheet={}, Row=[{}-{}], Col=[{}-{}], Size={}bytes, Type={}",
                                sheetIndex, row, row2, col, col2, imageData.data.length, imageData.mimeType);
                    }
                }
            }

            if (!sheetImages.isEmpty()) {
                extractedImages.put(sheetIndex, sheetImages);
                log.info("Sheet {} 提取了 {} 个位置的图片", sheetIndex, sheetImages.size());
            }
        }
    }

    /**
     * 提取XLS格式的图片
     */
    private void extractHSSFImages(HSSFWorkbook workbook) {
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            Map<String, List<ImageData>> sheetImages = new HashMap<>();

            HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
            if (patriarch == null) {
                log.debug("Sheet {} 没有Drawing对象", sheetIndex);
                continue;
            }

            List<HSSFShape> shapes = patriarch.getChildren();
            log.info("Sheet {} 找到 {} 个图形对象", sheetIndex, shapes.size());

            for (HSSFShape shape : shapes) {
                if (shape instanceof HSSFPicture) {
                    HSSFPicture picture = (HSSFPicture) shape;
                    HSSFClientAnchor anchor = (HSSFClientAnchor) picture.getAnchor();

                    if (anchor != null) {
                        int row = anchor.getRow1();
                        int col = anchor.getCol1();
                        int row2 = anchor.getRow2();
                        int col2 = anchor.getCol2();
                        String key = row + "_" + col;

                        ImageData imageData = new ImageData();
                        HSSFPictureData pictureData = picture.getPictureData();
                        imageData.data = pictureData.getData();
                        imageData.mimeType = pictureData.getMimeType();
                        imageData.pictureType = pictureData.getFormat();
                        imageData.row = row;
                        imageData.col = col;
                        imageData.row2 = row2;
                        imageData.col2 = col2;

                        sheetImages.computeIfAbsent(key, k -> new ArrayList<>()).add(imageData);

                        log.debug("提取图片: Sheet={}, Row=[{}-{}], Col=[{}-{}], Size={}bytes, Type={}",
                                sheetIndex, row, row2, col, col2, imageData.data.length, imageData.mimeType);
                    }
                }
            }

            if (!sheetImages.isEmpty()) {
                extractedImages.put(sheetIndex, sheetImages);
                log.info("Sheet {} 提取了 {} 个位置的图片", sheetIndex, sheetImages.size());
            }
        }
    }

    /**
     * 获取指定位置的图片
     *
     * @param sheetIndex sheet索引
     * @param row        行号
     * @param col        列号
     * @return 图片数据列表
     */
    public List<ImageData> getImages(int sheetIndex, int row, int col) {
        Map<String, List<ImageData>> sheetImages = extractedImages.get(sheetIndex);
        if (sheetImages == null) {
            return Collections.emptyList();
        }

        String key = row + "_" + col;
        return sheetImages.getOrDefault(key, Collections.emptyList());
    }

    /**
     * 获取第一张图片
     *
     * @param sheetIndex sheet索引
     * @param row        行号
     * @param col        列号
     * @return 图片数据，如果没有则返回null
     */
    public ImageData getFirstImage(int sheetIndex, int row, int col) {
        List<ImageData> images = getImages(sheetIndex, row, col);
        return images.isEmpty() ? null : images.get(0);
    }

    /**
     * 清理资源
     */
    public void clear() {
        extractedImages.clear();
    }

    /**
     * 获取覆盖指定单元格的图片（范围匹配）
     * <p>
     * 支持图片横跨多个单元格、遮挡单元格、超出行等情况
     * </p>
     *
     * @param sheetIndex sheet索引
     * @param row        行号
     * @param col        列号
     * @return 图片数据列表
     */
    public List<ImageData> getImagesCoveringCell(int sheetIndex, int row, int col) {
        Map<String, List<ImageData>> sheetImages = extractedImages.get(sheetIndex);
        if (sheetImages == null) {
            return Collections.emptyList();
        }

        List<ImageData> result = new ArrayList<>();

        // 遍历所有图片，检查是否覆盖目标单元格
        for (List<ImageData> images : sheetImages.values()) {
            for (ImageData image : images) {
                if (image.coversCell(row, col)) {
                    result.add(image);
                }
            }
        }

        // 按主锚点位置排序（左上角优先）
        result.sort((a, b) -> {
            if (a.row != b.row) return a.row - b.row;
            return a.col - b.col;
        });

        return result;
    }

    /**
     * 图片数据类
     */
    @Getter
    public static class ImageData {
        private byte[] data;
        private String mimeType;
        private int pictureType;
        private int row;      // 左上角行 (row1)
        private int col;      // 左上角列 (col1)
        private int row2;     // 右下角行
        private int col2;     // 右下角列
        private String fileName;

        /**
         * 检查图片是否覆盖指定单元格
         * <p>
         * 图片覆盖判定规则：
         * 1. 精确匹配：图片左上角在目标单元格
         * 2. 范围覆盖：图片跨越多个单元格，目标单元格在图片范围内
         * 3. 列匹配优先：对于图片导入，主要关注列位置
         * </p>
         *
         * @param targetRow 目标行
         * @param targetCol 目标列
         * @return 是否覆盖
         */
        public boolean coversCell(int targetRow, int targetCol) {
            // 精确匹配左上角
            if (row == targetRow && col == targetCol) {
                return true;
            }

            // 范围覆盖检查：图片区域包含目标单元格
            // row1 <= targetRow <= row2 且 col1 <= targetCol <= col2
            boolean rowInRange = row <= targetRow && targetRow <= row2;
            boolean colInRange = col <= targetCol && targetCol <= col2;

            return rowInRange && colInRange;
        }

        /**
         * 检查图片主锚点是否在指定行
         *
         * @param targetRow 目标行
         * @return 是否在指定行
         */
        public boolean isInRow(int targetRow) {
            // 图片主锚点在目标行，或图片跨越目标行
            return row == targetRow || (row <= targetRow && targetRow <= row2);
        }

        /**
         * 检查图片主锚点是否在指定列
         *
         * @param targetCol 目标列
         * @return 是否在指定列
         */
        public boolean isInColumn(int targetCol) {
            // 图片主锚点在目标列，或图片跨越目标列
            return col == targetCol || (col <= targetCol && targetCol <= col2);
        }

        /**
         * 转换为Base64格式
         */
        public String toBase64() {
            if (data == null || data.length == 0) {
                return null;
            }
            return "data:" + mimeType + ";base64," + Base64.getEncoder().encodeToString(data);
        }

        /**
         * 获取文件扩展名
         */
        public String getExtension() {
            if (mimeType == null) {
                return "png";
            }

            switch (mimeType.toLowerCase()) {
                case "image/jpeg":
                case "image/jpg":
                    return "jpg";
                case "image/png":
                    return "png";
                case "image/gif":
                    return "gif";
                case "image/bmp":
                    return "bmp";
                case "image/webp":
                    return "webp";
                default:
                    return "png";
            }
        }
    }
}