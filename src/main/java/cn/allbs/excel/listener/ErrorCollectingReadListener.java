package cn.allbs.excel.listener;

import cn.allbs.excel.model.ExcelReadError;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelDataConvertException;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * Error Collecting Read Listener
 * <p>
 * Collects all errors during Excel import instead of failing on first error.
 * Useful for batch imports where you want to see all validation issues at once.
 * </p>
 *
 * <p>Usage Examples:</p>
 * <pre>
 * // Create listener with error collection
 * ErrorCollectingReadListener&lt;UserDTO&gt; listener = new ErrorCollectingReadListener&lt;&gt;();
 * listener.setDataConsumer(users -&gt; userService.batchSave(users));
 *
 * // Read Excel
 * EasyExcel.read(file, UserDTO.class, listener).sheet().doRead();
 *
 * // Check for errors
 * if (listener.hasErrors()) {
 *     List&lt;ExcelReadError&gt; errors = listener.getErrors();
 *     // Display errors to user
 *     errors.forEach(error -&gt; System.out.println(error.getFormattedMessage()));
 * }
 *
 * // Get successfully read data
 * List&lt;UserDTO&gt; validData = listener.getValidData();
 * </pre>
 *
 * @param <T> Data type
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ErrorCollectingReadListener<T> extends AnalysisEventListener<T> {

	/**
	 * Batch size for data processing
	 */
	private static final int BATCH_SIZE = 500;

	/**
	 * Valid data list
	 */
	@Getter
	private final List<T> validData = new ArrayList<>();

	/**
	 * Error list
	 */
	@Getter
	private final List<ExcelReadError> errors = new ArrayList<>();

	/**
	 * Temporary batch data
	 */
	private final List<T> batchData = new ArrayList<>();

	/**
	 * Data consumer for batch processing
	 */
	private Consumer<List<T>> dataConsumer;

	/**
	 * Whether to continue on error
	 */
	private boolean continueOnError = true;

	/**
	 * Maximum errors allowed before stopping
	 */
	private int maxErrors = Integer.MAX_VALUE;

	/**
	 * Current sheet name
	 */
	private String currentSheetName;

	/**
	 * Current sheet index
	 */
	private Integer currentSheetIndex;

	/**
	 * Default constructor
	 */
	public ErrorCollectingReadListener() {
	}

	/**
	 * Constructor with data consumer
	 *
	 * @param dataConsumer Data consumer for batch processing
	 */
	public ErrorCollectingReadListener(Consumer<List<T>> dataConsumer) {
		this.dataConsumer = dataConsumer;
	}

	/**
	 * Set data consumer
	 *
	 * @param dataConsumer Data consumer
	 * @return this
	 */
	public ErrorCollectingReadListener<T> setDataConsumer(Consumer<List<T>> dataConsumer) {
		this.dataConsumer = dataConsumer;
		return this;
	}

	/**
	 * Set whether to continue on error
	 *
	 * @param continueOnError true to continue, false to stop
	 * @return this
	 */
	public ErrorCollectingReadListener<T> setContinueOnError(boolean continueOnError) {
		this.continueOnError = continueOnError;
		return this;
	}

	/**
	 * Set maximum errors allowed
	 *
	 * @param maxErrors Maximum errors
	 * @return this
	 */
	public ErrorCollectingReadListener<T> setMaxErrors(int maxErrors) {
		this.maxErrors = maxErrors;
		return this;
	}

	@Override
	public void invoke(T data, AnalysisContext context) {
		// Add valid data
		validData.add(data);
		batchData.add(data);

		// Process batch if size reached
		if (batchData.size() >= BATCH_SIZE) {
			processBatch();
		}
	}

	@Override
	public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
		currentSheetName = context.readSheetHolder().getSheetName();
		currentSheetIndex = context.readSheetHolder().getSheetNo();
		log.debug("Reading sheet: {} (index: {})", currentSheetName, currentSheetIndex);
	}

	@Override
	public void onException(Exception exception, AnalysisContext context) throws Exception {
		log.error("Error reading Excel at row {}", context.readRowHolder().getRowIndex(), exception);

		// Build error object
		ExcelReadError error = buildErrorFromException(exception, context);
		errors.add(error);

		// Check if should stop
		if (!continueOnError || errors.size() >= maxErrors) {
			log.error("Stopping Excel read due to error threshold. Total errors: {}", errors.size());
			throw exception;
		}

		// Continue reading
		log.debug("Continuing Excel read despite error. Total errors so far: {}", errors.size());
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		// Process remaining batch
		processBatch();

		log.info("Excel read completed. Valid rows: {}, Errors: {}", validData.size(), errors.size());

		if (!errors.isEmpty()) {
			log.warn("Excel read completed with {} errors:", errors.size());
			errors.forEach(error -> log.warn("  - {}", error.getFormattedMessage()));
		}
	}

	/**
	 * Process batch data
	 */
	private void processBatch() {
		if (batchData.isEmpty()) {
			return;
		}

		if (dataConsumer != null) {
			try {
				dataConsumer.accept(new ArrayList<>(batchData));
			}
			catch (Exception e) {
				log.error("Error processing batch data", e);
			}
		}

		batchData.clear();
	}

	/**
	 * Build error object from exception
	 */
	private ExcelReadError buildErrorFromException(Exception exception, AnalysisContext context) {
		ExcelReadError.ExcelReadErrorBuilder builder = ExcelReadError.builder().sheetName(currentSheetName)
			.sheetIndex(currentSheetIndex)
			.rowIndex(context.readRowHolder().getRowIndex())
			.excelRowNumber(context.readRowHolder().getRowIndex() + 1)
			.exception(exception);

		if (exception instanceof ExcelDataConvertException) {
			ExcelDataConvertException convertException = (ExcelDataConvertException) exception;

			builder.columnIndex(convertException.getColumnIndex())
				.errorType(ExcelReadError.ErrorType.TYPE_CONVERSION)
				.errorMessage("Data type conversion failed: " + exception.getMessage());

			// Try to get cell value
			try {
				Object cellData = convertException.getCellData();
				if (cellData != null) {
					builder.cellValue(cellData.toString());
				}
			}
			catch (Exception e) {
				// Ignore
			}
		}
		else {
			builder.errorType(ExcelReadError.ErrorType.UNKNOWN).errorMessage(exception.getMessage());
		}

		return builder.build();
	}

	/**
	 * Add a custom error
	 *
	 * @param error Error to add
	 */
	public void addError(ExcelReadError error) {
		errors.add(error);
	}

	/**
	 * Check if there are any errors
	 *
	 * @return true if there are errors, false otherwise
	 */
	public boolean hasErrors() {
		return !errors.isEmpty();
	}

	/**
	 * Get error count
	 *
	 * @return Number of errors
	 */
	public int getErrorCount() {
		return errors.size();
	}

	/**
	 * Get valid data count
	 *
	 * @return Number of valid rows
	 */
	public int getValidDataCount() {
		return validData.size();
	}

	/**
	 * Clear all data and errors
	 */
	public void clear() {
		validData.clear();
		errors.clear();
		batchData.clear();
	}

}
