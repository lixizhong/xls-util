package lxz.util.xls.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class XlsReader {
	
	private XlsReader(){}

	/**
	 * 读取excel到一个list
	 * @param filePath	excel文件名
	 * @param sheetIndex	sheet
	 * @param rowFrom	开始行。-1从第一个有效行开始
	 * @param rowTo		结束行。-1最后一个有效行
	 * @param cellList	单元格设置
	 * @param rowValueHandler	行转换处理器
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static <T> List<T> readXls(
			String filePath,
			int sheetIndex,
			int rowFrom,
			int rowTo,
			List<CellReaderSetting> cellList,
			RowReader<T> rowValueHandler
			) throws EncryptedDocumentException, InvalidFormatException, IOException {

		InputStream is = new FileInputStream(filePath);

		Workbook wb = WorkbookFactory.create(is);
		Sheet sheet = wb.getSheetAt(sheetIndex);

		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		
		rowFrom = rowFrom < 0 ? firstRowNum : rowFrom;
		rowTo = rowTo < 0 ? lastRowNum : rowTo;
		
		List<T> dataList = new ArrayList<T>(rowTo - rowFrom + 1);

		for (int rowIndex = rowFrom; rowIndex <= rowTo; rowIndex++) {
			Row row = sheet.getRow(rowIndex);

			Map<String, Object> cellMap = new HashMap<String, Object>(row.getLastCellNum() - row.getFirstCellNum() + 1);

			for (CellReaderSetting cs : cellList) {
				
				Cell cell = row.getCell(cs.getColumnIndex());
				
				if(cell == null){
					cellMap.put(cs.getPropName(), "");
					continue;
				}
				
				Object cellValue = null;
				int cellType = cell.getCellType();

				switch (cellType) {
	                case Cell.CELL_TYPE_NUMERIC:
	                    if (DateUtil.isCellDateFormatted(cell)) {
	                    	cellValue = cell.getDateCellValue();
	                    } else {
	                    	cellValue = cell.getNumericCellValue();
	                    }
	                    break;
	                default:
	                	cellValue = cell.getStringCellValue();
	            }
				
				cellMap.put(cs.getPropName(), cellValue);
			}

			T t = rowValueHandler.getRowValue(cellMap);

			dataList.add(t);
		}
		
		return dataList;
	}
}
