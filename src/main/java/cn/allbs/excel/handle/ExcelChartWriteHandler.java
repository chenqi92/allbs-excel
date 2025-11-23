package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelChart;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.Arrays;

/**
 * Excel Chart Write Handler
 * <p>
 * Creates charts in Excel sheets based on annotation configuration
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelChartWriteHandler implements WorkbookWriteHandler {

	private ExcelChart chartConfig;
	private Class<?> dataClass;
	private int dataStartRow;
	private int dataEndRow;

	/**
	 * Default constructor (for reflection instantiation)
	 */
	public ExcelChartWriteHandler() {
		this.chartConfig = null;
		this.dataClass = null;
		this.dataStartRow = 1;
		this.dataEndRow = 100;
	}

	public ExcelChartWriteHandler(ExcelChart chartConfig, Class<?> dataClass, int dataStartRow, int dataEndRow) {
		this.chartConfig = chartConfig;
		this.dataClass = dataClass;
		this.dataStartRow = dataStartRow;
		this.dataEndRow = dataEndRow;
	}

	@Override
	public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
		if (chartConfig == null || !chartConfig.enabled()) {
			log.debug("Chart not enabled or config is null");
			return;
		}

		log.info("=== ExcelChartWriteHandler.afterWorkbookDispose() called ===");
		log.info("Chart config: title='{}', type={}, xAxis='{}', yAxis={}",
			chartConfig.title(), chartConfig.type(), chartConfig.xAxisField(),
			Arrays.toString(chartConfig.yAxisFields()));

		try {
			Workbook workbook = writeWorkbookHolder.getWorkbook();
			log.info("Workbook has {} sheets", workbook.getNumberOfSheets());
			Sheet sheet = workbook.getSheetAt(0);
			log.info("Working with sheet: '{}', rows: {}", sheet.getSheetName(), sheet.getLastRowNum() + 1);

			if (sheet instanceof XSSFSheet) {
				createChart((XSSFSheet) sheet);
				log.info("Successfully created chart: {}", chartConfig.title());
			} else {
				log.warn("Sheet is not XSSFSheet, cannot create chart. Sheet class: {}", sheet.getClass().getName());
			}
		} catch (Exception e) {
			log.error("Failed to create chart: " + e.getMessage(), e);
		}
	}

	/**
	 * Create chart in sheet
	 */
	private void createChart(XSSFSheet sheet) {
		// Get or create drawing patriarch
		XSSFDrawing drawing = sheet.getDrawingPatriarch();
		if (drawing == null) {
			drawing = sheet.createDrawingPatriarch();
		}

		// Create anchor for chart position
		XSSFClientAnchor anchor = drawing.createAnchor(
			0, 0, 0, 0,
			chartConfig.startColumn(), chartConfig.startRow(),
			chartConfig.endColumn(), chartConfig.endRow()
		);

		// Create chart
		XSSFChart chart = drawing.createChart(anchor);

		// Set chart title
		if (!chartConfig.title().isEmpty()) {
			chart.setTitleText(chartConfig.title());
			chart.setTitleOverlay(false);
		}

		// Create chart based on type
		switch (chartConfig.type()) {
			case LINE:
				createLineChart(chart, sheet);
				break;
			case BAR:
				createBarChart(chart, sheet);
				break;
			case COLUMN:
				createColumnChart(chart, sheet);
				break;
			case PIE:
				createPieChart(chart, sheet);
				break;
			case AREA:
				createAreaChart(chart, sheet);
				break;
			case SCATTER:
				createScatterChart(chart, sheet);
				break;
			default:
				log.warn("Unsupported chart type: {}", chartConfig.type());
		}
	}

	/**
	 * Create line chart
	 */
	private void createLineChart(XSSFChart chart, XSSFSheet sheet) {
		// Create chart axes
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.crossAxis(bottomAxis);
		bottomAxis.crossAxis(leftAxis);

		// Set axis titles
		if (!chartConfig.xAxisTitle().isEmpty()) {
			bottomAxis.setTitle(chartConfig.xAxisTitle());
		}
		if (!chartConfig.yAxisTitle().isEmpty()) {
			leftAxis.setTitle(chartConfig.yAxisTitle());
		}

		// Create line chart data
		XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

		// Add data series
		addDataSeries(data, sheet);

		// Plot chart
		chart.plot(data);

		// Configure legend
		configureLegend(chart);
	}

	/**
	 * Create bar chart
	 */
	private void createBarChart(XSSFChart chart, XSSFSheet sheet) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.crossAxis(bottomAxis);
		bottomAxis.crossAxis(leftAxis);

		if (!chartConfig.xAxisTitle().isEmpty()) {
			bottomAxis.setTitle(chartConfig.xAxisTitle());
		}
		if (!chartConfig.yAxisTitle().isEmpty()) {
			leftAxis.setTitle(chartConfig.yAxisTitle());
		}

		XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		data.setBarDirection(BarDirection.BAR);

		addDataSeries(data, sheet);
		chart.plot(data);
		configureLegend(chart);
	}

	/**
	 * Create column chart
	 */
	private void createColumnChart(XSSFChart chart, XSSFSheet sheet) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.crossAxis(bottomAxis);
		bottomAxis.crossAxis(leftAxis);

		if (!chartConfig.xAxisTitle().isEmpty()) {
			bottomAxis.setTitle(chartConfig.xAxisTitle());
		}
		if (!chartConfig.yAxisTitle().isEmpty()) {
			leftAxis.setTitle(chartConfig.yAxisTitle());
		}

		XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		data.setBarDirection(BarDirection.COL);

		addDataSeries(data, sheet);
		chart.plot(data);
		configureLegend(chart);
	}

	/**
	 * Create pie chart
	 */
	private void createPieChart(XSSFChart chart, XSSFSheet sheet) {
		XDDFPieChartData data = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);

		// For pie chart, we only use the first Y-axis field
		if (chartConfig.yAxisFields().length > 0) {
			int xCol = findColumnIndex(sheet, chartConfig.xAxisField());
			int yCol = findColumnIndex(sheet, chartConfig.yAxisFields()[0]);

			if (xCol >= 0 && yCol >= 0) {
				// Try to determine if X-axis is numeric or string
				XDDFDataSource<?> categories = createCategoryDataSource(sheet, xCol);

				XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
					sheet,
					new CellRangeAddress(dataStartRow, dataEndRow, yCol, yCol)
				);

				XDDFPieChartData.Series series = (XDDFPieChartData.Series) data.addSeries(categories, values);
				series.setTitle(chartConfig.yAxisFields()[0], null);
				log.debug("Added pie chart series: {} (categories) vs {} (values)", chartConfig.xAxisField(), chartConfig.yAxisFields()[0]);
			}
		}

		chart.plot(data);
		configureLegend(chart);
	}

	/**
	 * Create area chart
	 */
	private void createAreaChart(XSSFChart chart, XSSFSheet sheet) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.crossAxis(bottomAxis);
		bottomAxis.crossAxis(leftAxis);

		if (!chartConfig.xAxisTitle().isEmpty()) {
			bottomAxis.setTitle(chartConfig.xAxisTitle());
		}
		if (!chartConfig.yAxisTitle().isEmpty()) {
			leftAxis.setTitle(chartConfig.yAxisTitle());
		}

		XDDFAreaChartData data = (XDDFAreaChartData) chart.createData(ChartTypes.AREA, bottomAxis, leftAxis);

		addDataSeries(data, sheet);
		chart.plot(data);
		configureLegend(chart);
	}

	/**
	 * Create scatter chart
	 */
	private void createScatterChart(XSSFChart chart, XSSFSheet sheet) {
		XDDFValueAxis bottomAxis = chart.createValueAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.crossAxis(bottomAxis);
		bottomAxis.crossAxis(leftAxis);

		if (!chartConfig.xAxisTitle().isEmpty()) {
			bottomAxis.setTitle(chartConfig.xAxisTitle());
		}
		if (!chartConfig.yAxisTitle().isEmpty()) {
			leftAxis.setTitle(chartConfig.yAxisTitle());
		}

		XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);

		// Set scatter style to show only markers (points) without lines
		data.setStyle(ScatterStyle.MARKER);

		// Scatter chart needs numeric data for both X and Y axes
		addScatterDataSeries(data, sheet);
		chart.plot(data);
		configureLegend(chart);
	}

	/**
	 * Add data series to chart
	 */
	private void addDataSeries(XDDFChartData data, XSSFSheet sheet) {
		int xCol = findColumnIndex(sheet, chartConfig.xAxisField());

		if (xCol < 0) {
			log.warn("X-axis field not found: {}", chartConfig.xAxisField());
			return;
		}

		// Create category data source
		XDDFDataSource<?> categories = XDDFDataSourcesFactory.fromStringCellRange(
			sheet,
			new CellRangeAddress(dataStartRow, dataEndRow, xCol, xCol)
		);

		// Add series for each Y-axis field
		for (String yField : chartConfig.yAxisFields()) {
			int yCol = findColumnIndex(sheet, yField);

			if (yCol >= 0) {
				XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
					sheet,
					new CellRangeAddress(dataStartRow, dataEndRow, yCol, yCol)
				);

				XDDFChartData.Series series = data.addSeries(categories, values);
				series.setTitle(yField, null);
			} else {
				log.warn("Y-axis field not found: {}", yField);
			}
		}
	}

	/**
	 * Add data series to scatter chart
	 * <p>
	 * Scatter charts require numeric data for both X and Y axes
	 * </p>
	 */
	private void addScatterDataSeries(XDDFScatterChartData data, XSSFSheet sheet) {
		int xCol = findColumnIndex(sheet, chartConfig.xAxisField());

		if (xCol < 0) {
			log.warn("X-axis field not found: {}", chartConfig.xAxisField());
			return;
		}

		// For scatter chart, X-axis should be numeric
		XDDFNumericalDataSource<Double> xValues = XDDFDataSourcesFactory.fromNumericCellRange(
			sheet,
			new CellRangeAddress(dataStartRow, dataEndRow, xCol, xCol)
		);

		// Add series for each Y-axis field
		for (String yField : chartConfig.yAxisFields()) {
			int yCol = findColumnIndex(sheet, yField);

			if (yCol >= 0) {
				XDDFNumericalDataSource<Double> yValues = XDDFDataSourcesFactory.fromNumericCellRange(
					sheet,
					new CellRangeAddress(dataStartRow, dataEndRow, yCol, yCol)
				);

				XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(xValues, yValues);
				series.setTitle(yField, null);
				log.debug("Added scatter series: {} vs {}", chartConfig.xAxisField(), yField);
			} else {
				log.warn("Y-axis field not found: {}", yField);
			}
		}
	}

	/**
	 * Configure chart legend
	 */
	private void configureLegend(XSSFChart chart) {
		if (chartConfig.showLegend()) {
			XDDFChartLegend legend = chart.getOrAddLegend();

			switch (chartConfig.legendPosition()) {
				case TOP:
					legend.setPosition(LegendPosition.TOP);
					break;
				case BOTTOM:
					legend.setPosition(LegendPosition.BOTTOM);
					break;
				case LEFT:
					legend.setPosition(LegendPosition.LEFT);
					break;
				case RIGHT:
					legend.setPosition(LegendPosition.RIGHT);
					break;
				case TOP_RIGHT:
					legend.setPosition(LegendPosition.TOP_RIGHT);
					break;
			}
		}
	}

	/**
	 * Create category data source with smart type detection
	 * <p>
	 * Tries to detect if the column contains numeric or string data
	 * </p>
	 */
	private XDDFDataSource<?> createCategoryDataSource(XSSFSheet sheet, int columnIndex) {
		CellRangeAddress range = new CellRangeAddress(dataStartRow, dataEndRow, columnIndex, columnIndex);

		// Check the first data cell to determine type
		Row firstDataRow = sheet.getRow(dataStartRow);
		if (firstDataRow != null) {
			Cell firstCell = firstDataRow.getCell(columnIndex);
			if (firstCell != null) {
				CellType cellType = firstCell.getCellType();

				// If it's a formula, get the cached formula result type
				if (cellType == CellType.FORMULA) {
					cellType = firstCell.getCachedFormulaResultType();
				}

				// If the cell is numeric, use numeric data source
				if (cellType == CellType.NUMERIC) {
					log.debug("Using numeric data source for column {}", columnIndex);
					return XDDFDataSourcesFactory.fromNumericCellRange(sheet, range);
				}
			}
		}

		// Default to string data source
		log.debug("Using string data source for column {}", columnIndex);
		return XDDFDataSourcesFactory.fromStringCellRange(sheet, range);
	}

	/**
	 * Find column index by field name
	 * <p>
	 * Supports both exact match and case-insensitive match
	 * </p>
	 */
	private int findColumnIndex(XSSFSheet sheet, String fieldName) {
		if (fieldName == null || fieldName.isEmpty()) {
			return -1;
		}

		Row headerRow = sheet.getRow(0);
		if (headerRow == null) {
			return -1;
		}

		// First try exact match
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null && fieldName.equals(cell.getStringCellValue())) {
				return i;
			}
		}

		// If not found, try case-insensitive match
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null && fieldName.equalsIgnoreCase(cell.getStringCellValue())) {
				log.debug("Found column '{}' using case-insensitive match: '{}'", fieldName, cell.getStringCellValue());
				return i;
			}
		}

		return -1;
	}
}
