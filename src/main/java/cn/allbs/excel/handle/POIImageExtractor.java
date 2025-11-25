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
                        String key = row + "_" + col;

                        ImageData imageData = new ImageData();
                        XSSFPictureData pictureData = picture.getPictureData();
                        imageData.data = pictureData.getData();
                        imageData.mimeType = pictureData.getMimeType();
                        imageData.pictureType = pictureData.getPictureType();
                        imageData.row = row;
                        imageData.col = col;
                        imageData.fileName = pictureData.getPackagePart() != null
                            ? pictureData.getPackagePart().getPartName().getName()
                            : "image_" + row + "_" + col;

                        sheetImages.computeIfAbsent(key, k -> new ArrayList<>()).add(imageData);

                        log.debug("提取图片: Sheet={}, Row={}, Col={}, Size={}bytes, Type={}",
                                sheetIndex, row, col, imageData.data.length, imageData.mimeType);
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
                        String key = row + "_" + col;

                        ImageData imageData = new ImageData();
                        HSSFPictureData pictureData = picture.getPictureData();
                        imageData.data = pictureData.getData();
                        imageData.mimeType = pictureData.getMimeType();
                        imageData.pictureType = pictureData.getFormat();
                        imageData.row = row;
                        imageData.col = col;

                        sheetImages.computeIfAbsent(key, k -> new ArrayList<>()).add(imageData);

                        log.debug("提取图片: Sheet={}, Row={}, Col={}, Size={}bytes, Type={}",
                                sheetIndex, row, col, imageData.data.length, imageData.mimeType);
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
     * 图片数据类
     */
    @Getter
    public static class ImageData {
        private byte[] data;
        private String mimeType;
        private int pictureType;
        private int row;
        private int col;
        private String fileName;

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