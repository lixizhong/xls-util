package lxz.util.xls.writer;

import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;


public class XlsWriter{
	
	private XlsWriter(){}
	
	/**
	 * 生成xls
	 * @param dataList
	 * @param setting
	 * @param os
	 * @throws Exception
	 */
	public static <T> void createXls(List<T> dataList, String[] appendMessages, XlsWriterSetting setting, OutputStream os) throws Exception{
		Workbook wb = new HSSFWorkbook();
		String safeName = WorkbookUtil.createSafeSheetName("Sheet1");
		Sheet sheet = wb.createSheet(safeName);
		
		List<CellWriterSetting> columnList = setting.getColumnList();
		
		if(columnList == null || columnList.size() == 0){
			wb.close();
			throw new IllegalArgumentException("生成xls文件出错，没有设置表头");
		}
		
		int rowIndex = setting.getStartRow();
		
		int columnSize = columnList.size();
		
		//生成标题，合并单元格，标题居中
		if(StringUtils.isNotBlank(setting.getTitle())){
			Row row = sheet.createRow(rowIndex++);
			Cell cell = row.createCell(0);
			
			CellStyle cellStyle = wb.createCellStyle();
	        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
	        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
	        cell.setCellStyle(cellStyle);
	        
	        cell.setCellValue(setting.getTitle());
	        
	        if(setting.isCreateIndex()){
	        	columnSize++;
			}
			
			sheet.addMergedRegion(new CellRangeAddress(
		            0, //first row (0-based)
		            0, //last row  (0-based)
		            0, //first column (0-based)
		            columnSize-1  //last column  (0-based)
		    ));
		}
		
		//单元格边框设置
		CellStyle cellBorderStyle = wb.createCellStyle();
	    cellBorderStyle.setBorderBottom(CellStyle.BORDER_THIN);
	    cellBorderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	    cellBorderStyle.setBorderLeft(CellStyle.BORDER_THIN);
	    cellBorderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	    cellBorderStyle.setBorderRight(CellStyle.BORDER_THIN);
	    cellBorderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
	    cellBorderStyle.setBorderTop(CellStyle.BORDER_THIN);
	    cellBorderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
	    
	    //默认时间格式设置
	    CreationHelper createHelper = wb.getCreationHelper();
	    CellStyle cellDateStyle = wb.createCellStyle();
	    cellDateStyle.cloneStyleFrom(cellBorderStyle);
	    cellDateStyle.setDataFormat(createHelper.createDataFormat().getFormat(setting.getDefaultDateFormat()));
	    
	    CellStyle cellDoubleStyle = wb.createCellStyle();
	    cellDoubleStyle.cloneStyleFrom(cellBorderStyle);
	    cellDoubleStyle.setDataFormat(createHelper.createDataFormat().getFormat("0.00"));
		
		//生成表头
		if(setting.isCreateHeader()){
			Row row = sheet.createRow(rowIndex++);
			
			int cellIndex = 0;
			
			if(setting.isCreateIndex()){
				Cell cell = row.createCell(cellIndex++);
				cell.setCellStyle(cellBorderStyle);
				cell.setCellValue(setting.getIndexName());
			}
			
			for (CellWriterSetting cellSetting : columnList) {
				Cell cell = row.createCell(cellIndex++);
				cell.setCellStyle(cellBorderStyle);
				cell.setCellValue(cellSetting.getColumnName());
			}
		}
		
		//填充表格数据
		int i = 0;
		for (T obj : dataList) {
			
			if(obj == null){
				continue;
			}
			
			Row row = sheet.createRow(rowIndex++);
			
			int cellIndex = 0;
			
			if(setting.isCreateIndex()){
				Cell cell = row.createCell(cellIndex++);
				cell.setCellStyle(cellBorderStyle);
				cell.setCellValue(i+1);
			}
			
			for (int j=0; j<columnList.size(); j++) {
				
				CellWriterSetting cs = columnList.get(j);
				
				Cell cell = row.createCell(cellIndex++);
				cell.setCellStyle(cellBorderStyle);
				Object propValue = null;
				
				String propName = cs.getPropName();
				CellWriter valueHandler = cs.getValueHandler();
				
				if(StringUtils.isBlank(propName)){
					//如果propName为空，那么就使用CellWriter设置表格的值，obj作为CellWriter的参数
					if(valueHandler != null){
						propValue = valueHandler.getCellValue(obj);
					}
				}else{
					try {
						propValue = PropertyUtils.getProperty(obj, propName);
					} catch (IllegalAccessException e) {
						e.printStackTrace();
						throw new IllegalArgumentException("生成xls文件出错，获取["+propName+"]出错");
					} catch (InvocationTargetException e) {
						e.printStackTrace();
						throw new IllegalArgumentException("生成xls文件出错，获取["+propName+"]出错");
					} catch (NoSuchMethodException e) {
						e.printStackTrace();
						throw new IllegalArgumentException("生成xls文件出错，获取["+propName+"]出错");
					} finally {
						wb.close();
					}
					//允许使用CellWriter对取到的值再做一次加工，CellWriter的参数为propValue
					if(valueHandler != null){
						propValue = valueHandler.getCellValue(propValue);
					}
				}
				
				if(propValue == null){
					continue;
				}
				
				setCellValue(cellDateStyle, cellDoubleStyle, cell, propValue);
			}
			i++;
		}
		
		for (int j = 0; j < columnSize; j++) {
			sheet.autoSizeColumn(j);
		}
		
		if(appendMessages != null && appendMessages.length > 0){
			for (String message : appendMessages) {
				Row row = sheet.createRow(rowIndex++);
				
				Cell cell = row.createCell(0);
				
				CellStyle cellStyle = wb.createCellStyle();
		        cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		        cell.setCellStyle(cellStyle);
		        
		        cell.setCellValue(message);
				
				sheet.addMergedRegion(new CellRangeAddress(
						rowIndex-1, //first row (0-based)
						rowIndex-1, //last row  (0-based)
			            0, //first column (0-based)
			            columnSize-1  //last column  (0-based)
			    ));
			}
		}
		
		wb.write(os);
		wb.close();
	}

	private static void setCellValue(CellStyle cellDateStyle, CellStyle cellDoubleStyle, Cell cell, Object propValue) {
		if(propValue instanceof Number){
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			
			if(propValue instanceof Integer){
				cell.setCellValue((Integer) propValue);
			} else if(propValue instanceof Double){
				cell.setCellStyle(cellDoubleStyle);
				cell.setCellValue((Double) propValue);
			} else if(propValue instanceof Float){
				cell.setCellStyle(cellDoubleStyle);
				cell.setCellValue((Float) propValue);
			} else if (propValue instanceof Long) {
				cell.setCellValue((Long) propValue);
			} else if(propValue instanceof BigDecimal){
				cell.setCellStyle(cellDoubleStyle);
				cell.setCellValue(((BigDecimal) propValue).doubleValue());
			} else {
				cell.setCellValue(propValue.toString());
			}
		} else if(propValue instanceof Date) {
			cell.setCellStyle(cellDateStyle);
			cell.setCellValue((Date) propValue);
		} else if(propValue instanceof Calendar) {
			cell.setCellStyle(cellDateStyle);
			cell.setCellValue((Calendar) propValue);
		} else {
			cell.setCellValue(propValue.toString());
		}
	}
	
}

