import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import lxz.util.xls.reader.CellReaderSetting;
import lxz.util.xls.reader.RowReader;
import lxz.util.xls.reader.XlsReader;
import lxz.util.xls.writer.CellWriterSetting;
import lxz.util.xls.writer.XlsWriter;
import lxz.util.xls.writer.XlsWriterSetting;

/**
 * @author lixizhong
 */
public class Test {

	public static void main(String[] args) throws Exception{
		
		System.out.println("开始读取");
		List<Student> dataList = testRead();
		System.out.println("共"+dataList.size()+"行");

		for (Student student : dataList) {
			System.out.println("姓名："+student.getName());
			System.out.println("性别："+student.getSex());
			System.out.println("年龄："+student.getAge());
			System.out.println("成绩："+student.getScore());
			System.out.println("生日："+student.getDate());
			
			System.out.println("=============================");
		}
		
		System.out.println("开始写入");
		
		testWrite(dataList);
		
		System.out.println("写入完毕");
	}
	
	
	public static void testWrite(List<Student> dataList) throws Exception{
		
		List<CellWriterSetting> columnList = new ArrayList<CellWriterSetting>();
		
		columnList.add(new CellWriterSetting("姓名", "name", null));
		columnList.add(new CellWriterSetting("性别", "sex", null));
		columnList.add(new CellWriterSetting("年龄", "age", null));
		columnList.add(new CellWriterSetting("分数", "score", null));
		columnList.add(new CellWriterSetting("生日", "date", null));
		
		XlsWriterSetting setting = new XlsWriterSetting(true, "序号", "拷贝的表格", true, true, 0, "yyyy-MM-dd", columnList);
		
		XlsWriter.createXls(dataList, null, setting, new FileOutputStream("拷贝的表格.xls"));
	}
	
	public static List<Student> testRead() throws Exception{
		List<CellReaderSetting> cellList = new ArrayList<CellReaderSetting>();
		
		cellList.add(new CellReaderSetting(0, "index"));
		cellList.add(new CellReaderSetting(1, "name"));
		cellList.add(new CellReaderSetting(2, "sex"));
		cellList.add(new CellReaderSetting(3, "age"));
		cellList.add(new CellReaderSetting(4, "score"));
		cellList.add(new CellReaderSetting(5, "date"));
		
		List<Student> dataList = XlsReader.readXls("读.xls", 0, 1, -1, cellList, new RowReader<Student>(){
			public Student getRowValue(Map<String, Object> cells) {
				Student s = new Student();
				
				String name = cells.get("name").toString();
				String sex = cells.get("sex").toString();
				int age = ((Double) cells.get("age")).intValue();
				Float score = ((Double)cells.get("score")).floatValue();
				
				Date date = (Date) cells.get("date");
				
				s.setName(name);
				s.setSex(sex);
				s.setAge(age);
				s.setScore(score);
				s.setDate(date);
				
				return s;
			}
		});
		
		return dataList;
	}
}
