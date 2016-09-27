package lxz.util.xls.reader;

public class CellReaderSetting{
	
	private final int columnIndex;
	private final String propName;
	
	public CellReaderSetting(int columnIndex, String propName) {
		super();
		this.columnIndex = columnIndex;
		this.propName = propName;
	}
	
	public String getPropName() {
		return propName;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

}