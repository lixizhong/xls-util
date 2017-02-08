package lxz.util.xls.writer;


public class CellWriterSetting{
	private final String columnName;
	private final String propName;
	private final CellWriter valueHandler;
	
	public CellWriterSetting(String columnName, String propName, CellWriter valueHandler) {
		super();
		this.columnName = columnName;
		this.propName = propName;
		this.valueHandler = valueHandler;
	}
	public String getColumnName() {
		return columnName;
	}
	public String getPropName() {
		return propName;
	}
	public CellWriter getValueHandler() {
		return valueHandler;
	}
}