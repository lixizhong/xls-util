package lxz.util.xls.writer;

 /**
 * 对于需要个性化显示的字段，需要实现此接口，否则只调用toString方法。 
 * @author lixizhong
 *
 */
public interface CellWriter{
	public Object getCellValue(Object obj);
}
