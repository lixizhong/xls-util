package lxz.util.xls.reader;

import java.util.Map;

public interface RowReader<T> {
	/**
	 * 从表示一行数据的map转换成一个目标对象，如果目标对象是数字类型，则map中对应的值为double类型
	 * map的value仅可能为java.util.Date, double, String三种类型
	 * @param cells
	 * @return
	 */
	public T getRowValue(Map<String, Object> cells);
}
