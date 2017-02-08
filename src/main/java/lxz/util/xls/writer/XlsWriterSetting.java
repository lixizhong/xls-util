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
	 * 默认日期显示格式,默认yyyy-MM-dd HH:mm:ss
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
	 * 从第几行开始输出（开始空几行）,不能小于0
	 */
	private final int startRow;
	/**
	 * 表头设置
	 */
	private List<CellWriterSetting> columnList;
    /**
     * 生成excel格式：xls或者xlsx, 默认xls
     */
    private XlsType type;
    /**
     * 序号开始值，默认1
     */
    private int indexStartValue;
    private String sheetName;
    /**
     * 使用流式写入，节省内存
     */
    private boolean useStreamWriter;

	public XlsWriterSetting(
			boolean createIndex, String indexName, 
			String title, boolean drawBorder, 
			boolean createHeader, int startRow,
			String defaultDateFormat, List<CellWriterSetting> columnList) {
		this.createIndex = createIndex;
		this.indexName = indexName;
		this.title = title;
		this.drawBorder = drawBorder;
		this.createHeader = createHeader;
		this.startRow = startRow;
		this.columnList = columnList;
		this.defaultDateFormat = defaultDateFormat;
        this.type = XlsType.XLS;
        this.indexStartValue = 1;
        this.sheetName = "Sheet1";
        if(startRow < 0) {
            throw new IllegalArgumentException("开始行不能小于0");
        }
        if(columnList == null || columnList.size() == 0){
            throw new IllegalArgumentException("生成excel文件出错，没有设置表头");
        }
        this.useStreamWriter = false;
    }

    public XlsWriterSetting(
            boolean createIndex, String indexName,
            String title, boolean drawBorder,
            boolean createHeader, int startRow,
            String defaultDateFormat, List<CellWriterSetting> columnList,
            XlsType type, int indexStartValue) {
        this.createIndex = createIndex;
        this.indexName = indexName;
        this.title = title;
        this.drawBorder = drawBorder;
        this.createHeader = createHeader;
        this.startRow = startRow;
        this.columnList = columnList;
        this.defaultDateFormat = defaultDateFormat;
        this.type = type;
        this.indexStartValue = indexStartValue;
        this.sheetName = "Sheet1";
        if(startRow < 0) {
            throw new IllegalArgumentException("开始行不能小于0");
        }
        if(columnList == null || columnList.size() == 0){
            throw new IllegalArgumentException("生成excel文件出错，没有设置表头");
        }
        this.useStreamWriter = false;
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
        if(defaultDateFormat == null || defaultDateFormat.trim().length() == 0) {
            return "yyyy-MM-dd HH:mm:ss";
        }
		return defaultDateFormat;
	}

    public XlsType getType() {
        return type;
    }

    public void setType(XlsType type) {
        this.type = type;
    }

    public int getIndexStartValue() {
        return indexStartValue;
    }

    public void setIndexStartValue(int indexStartValue) {
        this.indexStartValue = indexStartValue;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setColumnList(List<CellWriterSetting> columnList) {
        this.columnList = columnList;
    }

    public boolean isUseStreamWriter() {
        return useStreamWriter;
    }

    public void setUseStreamWriter(boolean useStreamWriter) {
        this.useStreamWriter = useStreamWriter;
    }
}
