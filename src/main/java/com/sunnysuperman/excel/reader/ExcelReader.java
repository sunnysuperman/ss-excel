package com.sunnysuperman.excel.reader;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sunnysuperman.commons.util.FileUtil;
import com.sunnysuperman.commons.util.StringUtil;
import com.sunnysuperman.excel.ExcelException;
import com.sunnysuperman.excel.ExcelUtils;

public class ExcelReader {
	private static final Logger LOG = LoggerFactory.getLogger(ExcelReader.class);

	private File file; // 读取文件
	private boolean streaming; // 是否用流形式
	private int streamingRowCacheSize = 100; // 流式读取条数
	private Sheet sheet; // 读取表格
	private Handler handler; // 数据处理器
	private ExcelColumn[] columns; // 列
	private boolean columnsInOrder; // 列顺序是否需要保持一致
	private boolean firstRowNumAsOne; // 首行行号是否以1开始
	private boolean copy; // 读取的同时拷贝到另一个表格里
	private int copyRowCacheSize; // 拷贝行缓存条数
	private Sheet copySheet; // 指定拷贝表格

	int[] columnIndexes;
	Workbook workbook; // 原始工作簿
	Workbook copyWorkbook; // 拷贝工作簿

	public ExcelReader(File file) {
		super();
		this.file = file;
	}

	public ExcelReader(Sheet sheet) {
		super();
		this.sheet = sheet;
	}

	private void loadSheet() throws ExcelException {
		if (sheet != null) {
			return;
		}
		if (streaming && streamingRowCacheSize <= 0) {
			throw new IllegalArgumentException("streamingRowCacheSize");
		}
		workbook = ExcelUtils.loadWorkbook(file, streaming ? streamingRowCacheSize : 0);
		sheet = ExcelUtils.ensureSheet(workbook, 0);

		if (copy && copySheet == null) {
			copyWorkbook = copyRowCacheSize == 0 ? ExcelUtils.newWorkbook() : ExcelUtils.newWorkbook(copyRowCacheSize);
			copySheet = ExcelUtils.ensureSheet(copyWorkbook, 0);
		}
	}

	public void read() throws ExcelException, HandlerException {
		try {
			// 准备表格
			loadSheet();
			// 读取
			doRead();
		} catch (Exception e) {
			LOG.error(null, e);
			// 如果出错了，关闭自动生成的拷贝工作簿，否则不关闭（另有用途，如写入到文件等）
			if (copyWorkbook != null) {
				FileUtil.close(copyWorkbook);
			}
		} finally {
			// 原始工作簿，需要关闭（如果是传入sheet的不关闭，交由调用方自己关闭）
			if (workbook != null && file != null) {
				FileUtil.close(workbook);
			}
		}
	}

	private void doRead() throws ExcelException, HandlerException {
		// 开始
		int firstRow = streaming ? 0 : sheet.getFirstRowNum();
		int lastRow = streaming ? -1 : sheet.getLastRowNum();
		boolean ok = handler.onStart(this, lastRow - firstRow);
		if (!ok) {
			return;
		}
		// 逐行读取数据
		if (streaming) {
			for (Row row : sheet) {
				if (row != null) {
					if (columnIndexes == null) {
						columnIndexes = readHeader(row);
					} else {
						ok = readRow(row);
					}
					if (!ok) {
						break;
					}
				}
			}
		} else {
			columnIndexes = readHeader(sheet.getRow(firstRow));
			for (int i = firstRow + 1; i <= lastRow; i++) {
				Row row = sheet.getRow(i);
				if (row != null) {
					ok = readRow(row);
					if (!ok) {
						break;
					}
				}
			}
		}
		// 结束
		handler.onEnd(this);
	}

	private int[] readHeader(Row row) throws ExcelException, HandlerException {
		int[] indexes = new int[columns.length];
		if (columnsInOrder) {
			for (int i = 0; i < columns.length; i++) {
				Cell cell = row.getCell(i);
				String cellValue = null;
				if (cell != null) {
					try {
						cellValue = cell.getStringCellValue();
					} catch (Exception ex) {
						// ignore
					}
				}
				if (cellValue == null || !cellValue.equals(columns[i].getTitle())) {
					throw new ExcelException(ExcelException.ERROR_COLUMN_NOT_MATCH, i);
				}
				indexes[i] = i;
			}
		} else {
			for (int i = 0; i < columns.length; i++) {
				int index = findCellIndex(row, columns[i]);
				if (index < 0) {
					throw new ExcelException(ExcelException.ERROR_COULD_NOT_FIND_COLUMN, i);
				}
				indexes[i] = index;
			}
		}
		if (copySheet != null) {
			copyRow(row, true);
		}
		handler.onHeaderRead(this, row);
		return indexes;
	}

	private boolean readRow(Row row) throws HandlerException {
		if (row == null) {
			return true;
		}
		if (copySheet != null) {
			copyRow(row, false);
		}
		Map<String, Object> data = new HashMap<>();
		List<ExcelColumn> errorColumns = readRow(data, row);
		return handler.onData(this, data, firstRowNumAsOne ? row.getRowNum() + 1 : row.getRowNum(), errorColumns);
	}

	private void copyRow(Row row, boolean isHeader) throws HandlerException {
		Row copyRow = copySheet.createRow(row.getRowNum());
		ExcelUtils.copyRow(row, copyRow);
		handler.onRowCopied(this, row, copyRow, isHeader);
	}

	private List<ExcelColumn> readRow(Map<String, Object> data, Row row) {
		List<ExcelColumn> errorColumns = null;
		for (int k = 0; k < columnIndexes.length; k++) {
			int index = columnIndexes[k];
			if (index < 0) {
				continue;
			}
			Cell cell = row.getCell(index);
			if (cell == null) {
				continue;
			}
			ExcelColumn column = columns[k];
			Object value;
			try {
				value = ExcelUtils.getCellValue(cell, column.getType());
			} catch (Exception ex) {
				try {
					value = ExcelUtils.getCellValue(cell);
				} catch (Exception ex2) {
					value = null;
				}
				if (errorColumns == null) {
					errorColumns = new ArrayList<>(columns.length);
				}
				errorColumns.add(column);
			}
			data.put(column.getKey(), value);
		}
		return errorColumns;
	}

	private int findCellIndex(Row row, ExcelColumn column) throws ExcelException {
		int firstCell = row.getFirstCellNum();
		int lastCell = row.getLastCellNum();
		for (int i = firstCell; i <= lastCell; i++) {
			Cell cell = row.getCell(i);
			if (cell == null) {
				continue;
			}
			String title = StringUtil.trimToEmpty(ExcelUtils.getStringCellValue(cell));
			switch (column.getMatchMode()) {
			case EXACTLY:
				if (column.getTitle().equals(title)) {
					return i;
				}
				break;
			case STARTS_WITH:
				/* XX(*) startsWith XX */
				if (column.getTitle().startsWith(title)) {
					return i;
				}
				break;
			case FUZZY:
				/* XY indexOf XXY */
				if (column.getTitle().indexOf(title) >= 0) {
					return i;
				}
				break;
			default:
				break;
			}
		}
		return -1;
	}

	public File getFile() {
		return file;
	}

	public ExcelReader setFile(File file) {
		this.file = file;
		return this;
	}

	public boolean isStreaming() {
		return streaming;
	}

	public ExcelReader setStreaming(boolean streaming) {
		this.streaming = streaming;
		return this;
	}

	public int getStreamingRowCacheSize() {
		return streamingRowCacheSize;
	}

	public ExcelReader setStreamingRowCacheSize(int streamingRowCacheSize) {
		this.streamingRowCacheSize = streamingRowCacheSize;
		return this;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public ExcelReader setSheet(Sheet sheet) {
		this.sheet = sheet;
		return this;
	}

	public Handler getHandler() {
		return handler;
	}

	public ExcelReader setHandler(Handler handler) {
		this.handler = handler;
		return this;
	}

	public ExcelColumn[] getColumns() {
		return columns;
	}

	public ExcelReader setColumns(ExcelColumn[] columns) {
		this.columns = columns;
		return this;
	}

	public boolean isColumnsInOrder() {
		return columnsInOrder;
	}

	public ExcelReader setColumnsInOrder(boolean columnsInOrder) {
		this.columnsInOrder = columnsInOrder;
		return this;
	}

	public boolean isFirstRowNumAsOne() {
		return firstRowNumAsOne;
	}

	public ExcelReader setFirstRowNumAsOne(boolean firstRowNumAsOne) {
		this.firstRowNumAsOne = firstRowNumAsOne;
		return this;
	}

	public boolean isCopy() {
		return copy;
	}

	public ExcelReader setCopy(boolean copy) {
		this.copy = copy;
		return this;
	}

	public int getCopyRowCacheSize() {
		return copyRowCacheSize;
	}

	public ExcelReader setCopyRowCacheSize(int copyRowCacheSize) {
		this.copyRowCacheSize = copyRowCacheSize;
		this.copy = true;
		return this;
	}

	public Sheet getCopySheet() {
		return copySheet;
	}

	public ExcelReader setCopySheet(Sheet copySheet) {
		this.copySheet = copySheet;
		this.copy = true;
		return this;
	}

}