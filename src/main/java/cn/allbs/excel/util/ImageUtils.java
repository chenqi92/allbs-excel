package cn.allbs.excel.util;

import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.UUID;

/**
 * 图片工具类
 * <p>
 * 提供图片的加载、保存、转换等功能，支持图片导入导出
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-21
 */
@Slf4j
public class ImageUtils {

	/**
	 * 将图片URL转换为Base64字符串
	 *
	 * @param imageUrl 图片URL
	 * @return Base64字符串（带data:image/xxx;base64,前缀）
	 */
	public static String urlToBase64(String imageUrl) {
		try (InputStream is = new URL(imageUrl).openStream(); ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

			byte[] buffer = new byte[4096];
			int bytesRead;
			while ((bytesRead = is.read(buffer)) != -1) {
				baos.write(buffer, 0, bytesRead);
			}

			byte[] imageBytes = baos.toByteArray();
			String base64 = Base64.getEncoder().encodeToString(imageBytes);
			String format = detectImageFormat(imageBytes);

			return String.format("data:image/%s;base64,%s", format, base64);
		}
		catch (Exception e) {
			log.error("Failed to load image from URL: {}", imageUrl, e);
			return null;
		}
	}

	/**
	 * 将本地图片文件转换为Base64字符串
	 *
	 * @param filePath 本地文件路径
	 * @return Base64字符串（带data:image/xxx;base64,前缀）
	 */
	public static String fileToBase64(String filePath) {
		try {
			Path path = Paths.get(filePath);
			byte[] imageBytes = Files.readAllBytes(path);
			String base64 = Base64.getEncoder().encodeToString(imageBytes);
			String format = detectImageFormat(imageBytes);

			return String.format("data:image/%s;base64,%s", format, base64);
		}
		catch (Exception e) {
			log.error("Failed to load image from file: {}", filePath, e);
			return null;
		}
	}

	/**
	 * 将字节数组转换为Base64字符串
	 *
	 * @param imageBytes 图片字节数组
	 * @return Base64字符串（带data:image/xxx;base64,前缀）
	 */
	public static String bytesToBase64(byte[] imageBytes) {
		if (imageBytes == null || imageBytes.length == 0) {
			return null;
		}

		String base64 = Base64.getEncoder().encodeToString(imageBytes);
		String format = detectImageFormat(imageBytes);

		return String.format("data:image/%s;base64,%s", format, base64);
	}

	/**
	 * 将Base64字符串转换为字节数组
	 *
	 * @param base64Str Base64字符串（可以带或不带data:image/xxx;base64,前缀）
	 * @return 图片字节数组
	 */
	public static byte[] base64ToBytes(String base64Str) {
		if (base64Str == null || base64Str.isEmpty()) {
			return null;
		}

		// 移除前缀
		String base64Data = base64Str;
		if (base64Str.contains(",")) {
			base64Data = base64Str.substring(base64Str.indexOf(",") + 1);
		}

		try {
			return Base64.getDecoder().decode(base64Data);
		}
		catch (Exception e) {
			log.error("Failed to decode base64 string", e);
			return null;
		}
	}

	/**
	 * 将Base64字符串保存为本地文件
	 *
	 * @param base64Str  Base64字符串
	 * @param outputPath 输出文件路径
	 * @return 是否保存成功
	 */
	public static boolean base64ToFile(String base64Str, String outputPath) {
		byte[] imageBytes = base64ToBytes(base64Str);
		if (imageBytes == null) {
			return false;
		}

		try {
			Path path = Paths.get(outputPath);
			Files.createDirectories(path.getParent());
			Files.write(path, imageBytes);
			log.debug("Saved image to: {}", outputPath);
			return true;
		}
		catch (Exception e) {
			log.error("Failed to save image to file: {}", outputPath, e);
			return false;
		}
	}

	/**
	 * 检测图片格式
	 *
	 * @param imageBytes 图片字节数组
	 * @return 图片格式（png、jpg、gif等）
	 */
	public static String detectImageFormat(byte[] imageBytes) {
		if (imageBytes == null || imageBytes.length < 4) {
			return "png";
		}

		// PNG: 89 50 4E 47
		if (imageBytes[0] == (byte) 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E && imageBytes[3] == 0x47) {
			return "png";
		}

		// JPEG: FF D8 FF
		if (imageBytes[0] == (byte) 0xFF && imageBytes[1] == (byte) 0xD8 && imageBytes[2] == (byte) 0xFF) {
			return "jpeg";
		}

		// GIF: 47 49 46 38
		if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46 && imageBytes[3] == 0x38) {
			return "gif";
		}

		// BMP: 42 4D
		if (imageBytes[0] == 0x42 && imageBytes[1] == 0x4D) {
			return "bmp";
		}

		// WebP: 52 49 46 46 ... 57 45 42 50
		if (imageBytes.length >= 12 && imageBytes[0] == 0x52 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46
				&& imageBytes[3] == 0x46 && imageBytes[8] == 0x57 && imageBytes[9] == 0x45 && imageBytes[10] == 0x42
				&& imageBytes[11] == 0x50) {
			return "webp";
		}

		// 默认返回png
		return "png";
	}

	/**
	 * 获取图片的实际尺寸
	 *
	 * @param imageBytes 图片字节数组
	 * @return 图片尺寸数组 [width, height]，如果获取失败返回 null
	 */
	public static int[] getImageSize(byte[] imageBytes) {
		if (imageBytes == null || imageBytes.length == 0) {
			return null;
		}

		try (ByteArrayInputStream bais = new ByteArrayInputStream(imageBytes)) {
			BufferedImage image = ImageIO.read(bais);
			if (image != null) {
				return new int[] { image.getWidth(), image.getHeight() };
			}
		}
		catch (Exception e) {
			log.debug("Failed to get image size", e);
		}

		return null;
	}

	/**
	 * 生成随机文件名
	 *
	 * @param extension 文件扩展名（如 "png", "jpg"）
	 * @return 随机文件名
	 */
	public static String generateRandomFilename(String extension) {
		return UUID.randomUUID().toString().replace("-", "") + "." + extension;
	}

	/**
	 * 保存图片到临时目录
	 *
	 * @param base64Str Base64字符串
	 * @return 保存后的文件路径，失败返回null
	 */
	public static String saveToTempDirectory(String base64Str) {
		byte[] imageBytes = base64ToBytes(base64Str);
		if (imageBytes == null) {
			return null;
		}

		String format = detectImageFormat(imageBytes);
		String filename = generateRandomFilename(format);

		try {
			Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"), "excel-images");
			Files.createDirectories(tempDir);

			Path tempFile = tempDir.resolve(filename);
			Files.write(tempFile, imageBytes);

			log.debug("Saved image to temp directory: {}", tempFile);
			return tempFile.toString();
		}
		catch (Exception e) {
			log.error("Failed to save image to temp directory", e);
			return null;
		}
	}

}
