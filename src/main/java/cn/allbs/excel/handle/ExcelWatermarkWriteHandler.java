package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelWatermark;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;

/**
 * Excel Watermark Write Handler
 * <p>
 * Adds watermark to Excel sheets
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelWatermarkWriteHandler implements WorkbookWriteHandler {

	private final ExcelWatermark watermark;

	public ExcelWatermarkWriteHandler(ExcelWatermark watermark) {
		this.watermark = watermark;
	}

	@Override
	public void afterWorkbookCreate(WriteWorkbookHolder writeWorkbookHolder) {
		if (!watermark.enabled()) {
			return;
		}

		try {
			Workbook workbook = writeWorkbookHolder.getWorkbook();

			// Add watermark to all sheets
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				Sheet sheet = workbook.getSheetAt(i);
				addWatermarkToSheet((XSSFSheet) sheet);
			}

			log.info("Successfully added watermark to Excel: {}", watermark.text());
		} catch (Exception e) {
			log.error("Failed to add watermark to Excel", e);
		}
	}

	/**
	 * Add watermark to sheet
	 */
	private void addWatermarkToSheet(XSSFSheet sheet) {
		try {
			// Create watermark image
			BufferedImage watermarkImage = createWatermarkImage();

			// Convert image to byte array
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			ImageIO.write(watermarkImage, "PNG", baos);
			byte[] imageBytes = baos.toByteArray();

			// Add image to workbook
			XSSFWorkbook workbook = sheet.getWorkbook();
			int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);

			// Create drawing patriarch
			XSSFDrawing drawing = sheet.createDrawingPatriarch();

			// Create anchor
			XSSFClientAnchor anchor = new XSSFClientAnchor(
				0, 0, 0, 0,
				watermark.startColumn(), watermark.startRow(),
				watermark.startColumn() + 20, watermark.startRow() + 30
			);

			// Create picture
			XSSFPicture picture = drawing.createPicture(anchor, pictureIdx);

			// Resize picture to fit the watermark
			picture.resize();

			log.debug("Added watermark to sheet: {}", sheet.getSheetName());
		} catch (Exception e) {
			log.error("Failed to add watermark to sheet: {}", sheet.getSheetName(), e);
		}
	}

	/**
	 * Create watermark image
	 */
	private BufferedImage createWatermarkImage() {
		// Parse color
		Color color = parseColor(watermark.color());
		Color watermarkColor = new Color(
			color.getRed(),
			color.getGreen(),
			color.getBlue(),
			(int) (watermark.opacity() * 255)
		);

		// Create font
		Font font = new Font(watermark.fontName(), Font.BOLD, watermark.fontSize());

		// Calculate image size
		BufferedImage tempImage = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
		Graphics2D tempG2d = tempImage.createGraphics();
		tempG2d.setFont(font);
		FontMetrics fm = tempG2d.getFontMetrics();
		int textWidth = fm.stringWidth(processWatermarkText(watermark.text()));
		int textHeight = fm.getHeight();
		tempG2d.dispose();

		// Calculate rotated image size
		double radians = Math.toRadians(watermark.rotation());
		int imageWidth = (int) Math.abs(textWidth * Math.cos(radians)) + (int) Math.abs(textHeight * Math.sin(radians)) + 100;
		int imageHeight = (int) Math.abs(textWidth * Math.sin(radians)) + (int) Math.abs(textHeight * Math.cos(radians)) + 100;

		// Create image
		BufferedImage image = new BufferedImage(
			imageWidth * 3 + watermark.horizontalSpacing() * 2,
			imageHeight * 3 + watermark.verticalSpacing() * 2,
			BufferedImage.TYPE_INT_ARGB
		);

		Graphics2D g2d = image.createGraphics();

		// Set rendering hints for better quality
		g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
		g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

		// Set font and color
		g2d.setFont(font);
		g2d.setColor(watermarkColor);

		// Draw watermark text in a grid pattern
		String text = processWatermarkText(watermark.text());
		for (int row = 0; row < 3; row++) {
			for (int col = 0; col < 3; col++) {
				int x = col * (imageWidth + watermark.horizontalSpacing()) + imageWidth / 2;
				int y = row * (imageHeight + watermark.verticalSpacing()) + imageHeight / 2;

				// Save original transform
				AffineTransform originalTransform = g2d.getTransform();

				// Rotate around the center point
				g2d.rotate(radians, x, y);

				// Draw text centered
				FontMetrics metrics = g2d.getFontMetrics();
				int textX = x - metrics.stringWidth(text) / 2;
				int textY = y + metrics.getAscent() / 2;

				g2d.drawString(text, textX, textY);

				// Restore original transform
				g2d.setTransform(originalTransform);
			}
		}

		g2d.dispose();

		return image;
	}

	/**
	 * Process watermark text with placeholders
	 */
	private String processWatermarkText(String text) {
		String processed = text;

		// Replace date/time placeholders
		java.time.LocalDateTime now = java.time.LocalDateTime.now();
		java.time.format.DateTimeFormatter dateFormatter = java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd");
		java.time.format.DateTimeFormatter timeFormatter = java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss");
		java.time.format.DateTimeFormatter datetimeFormatter = java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

		processed = processed.replace("${date}", now.format(dateFormatter));
		processed = processed.replace("${time}", now.format(timeFormatter));
		processed = processed.replace("${datetime}", now.format(datetimeFormatter));

		// Note: ${user.name} would need to be resolved by Spring's SpEL evaluator in actual usage
		// For now, we'll use a placeholder
		processed = processed.replace("${user.name}", System.getProperty("user.name", "User"));

		return processed;
	}

	/**
	 * Parse color from hex string
	 */
	private Color parseColor(String hexColor) {
		try {
			String hex = hexColor.startsWith("#") ? hexColor.substring(1) : hexColor;
			int r = Integer.parseInt(hex.substring(0, 2), 16);
			int g = Integer.parseInt(hex.substring(2, 4), 16);
			int b = Integer.parseInt(hex.substring(4, 6), 16);
			return new Color(r, g, b);
		} catch (Exception e) {
			log.warn("Invalid color format: {}, using default gray", hexColor);
			return Color.LIGHT_GRAY;
		}
	}
}
