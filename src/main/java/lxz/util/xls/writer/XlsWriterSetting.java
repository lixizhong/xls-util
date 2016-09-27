package lxz.util.xls.writer;

import java.util.List;

public class XlsWriterSetting{
	/**
	 * 是否生成序列号
	 */
	private final boolean createIndex;
	/**
	 * 序列号表头名称
	 */
	private final String indexName;
	/**
	 * 默认日期显示格式
	 */
	private final String defaultDateFormat;
	/**
	 * 表标题
	 */
	private final String title;
	/**
	 * 是否画边框
	 */
	private final boolean drawBorder;
	/**
	 * 是否生成表头
	 */
	private final boolean createHeader;
	/**
	 * 从第几行开始输出（开始空几行）
	 */
	private final int startRow;
	/**
	 * 表头设置
	 */
	private List<CellWriterSetting> columnList;

	public XlsWriterSetting(
			boolean createIndex, String indexName, 
			String title, boolean drawBorder, 
			Boolean createHeader, int startRow,
			String defaultDateFormat, List<CellWriterSetting> columnList) {
		super();
		this.createIndex = createIndex;
		this.indexName = indexName;
		this.title = title;
		this.drawBorder = drawBorder;
		this.createHeader = createHeader;
		this.startRow = startRow;
		this.columnList = columnList;
		this.defaultDateFormat = defaultDateFormat;
	}

	public boolean isCreateIndex() {
		return createIndex;
	}

	public String getIndexName() {
		return indexName;
	}

	public String getTitle() {
		return title;
	}

	public boolean isDrawBorder() {
		return drawBorder;
	}

	public boolean isCreateHeader() {
		return createHeader;
	}

	public int getStartRow() {
		return startRow;
	}

	public List<CellWriterSetting> getColumnList() {
		return columnList;
	}

	public String getDefaultDateFormat() {
		return defaultDateFormat;
	}
}
