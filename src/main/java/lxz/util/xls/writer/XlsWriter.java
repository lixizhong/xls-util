package lxz.util.xls.writer;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;


public class XlsWriter{

    private XlsWriter(){}

    /**
     * 生成excel
     * 如果生存xls格式，要确保生成的excel表格总行数（表头+数据+空行）不能超过65536，xlsx格式支持1048576行,
     * OutputStream需用户自己关闭
     */
    public static <T> void createXls(List<T> dataList, String[] appendMessages, XlsWriterSetting setting, OutputStream os) throws Exception {
        ExcelFile excel = initExcelFile(setting);
        appendList(excel, dataList);
        appendMessage(excel, appendMessages);
        saveFile(excel, os);
        closeFile(excel);
    }

    public static void saveFile(ExcelFile excel, OutputStream os) throws Exception{
        excel.getWorkbook().write(os);
    }

    public static void closeFile(ExcelFile excel) throws Exception{
        excel.getWorkbook().close();
        if(excel.getWorkbook() instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) excel.getWorkbook()).dispose();
        }
    }

    public static void appendMessage(ExcelFile excel, String[] appendMessages) {
        if(appendMessages != null && appendMessages.length > 0){
            Sheet sheet = getSheet(excel);
            for (String message : appendMessages) {
                Row row = sheet.createRow(excel.getAndPlusCurrentRow());

                Cell cell = row.createCell(0);

                CellStyle cellStyle = excel.getWorkbook().createCellStyle();
                cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
                cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                cell.setCellStyle(cellStyle);

                cell.setCellValue(message);

                sheet.addMergedRegion(new CellRangeAddress(
                        excel.getCurrentRow()-1, //first row (0-based)
                        excel.getCurrentRow()-1, //last row  (0-based)
                        0, //first column (0-based)
                        excel.getColumnSize()-1  //last column  (0-based)
                ));
            }
        }
    }

    public static <T> void appendList(ExcelFile excel, List<T> dataList) {
        XlsWriterSetting setting = excel.getSetting();
        Workbook wb = excel.getWorkbook();
        //单元格边框设置
        CellStyle cellBorderStyle = excel.getCellBorderStyle();
        //默认时间格式设置
        CellStyle cellDateStyle = excel.getCellDateStyle();
        //浮点数格式
        CellStyle cellDoubleStyle = excel.getCellDoubleStyle();
        //整数格式
        CellStyle cellIntStyle = excel.getCellIntStyle();

        Sheet sheet = wb.getSheet(WorkbookUtil.createSafeSheetName(setting.getSheetName()));
        List<CellWriterSetting> columnList = excel.getSetting().getColumnList();

        //填充表格数据
        for (T obj : dataList) {
            if(obj == null){
                continue;
            }

            Row row = sheet.createRow(excel.getAndPlusCurrentRow());

            int cellIndex = 0;

            if(setting.isCreateIndex()){
                Cell cell = row.createCell(cellIndex++);
                cell.setCellStyle(cellBorderStyle);
                cell.setCellValue(excel.getCurrentAndPlusIndexNum());
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
                    } catch (Exception e) {
                        try {
                            wb.close();
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        e.printStackTrace();
                        throw new IllegalArgumentException("生成excel文件出错，获取["+propName+"]出错");
                    }
                    //允许使用CellWriter对取到的值再做一次加工，CellWriter的参数为propValue
                    if(valueHandler != null){
                        propValue = valueHandler.getCellValue(propValue);
                    }
                }

                if(propValue == null){
                    continue;
                }

                setCellValue(cellDateStyle, cellDoubleStyle, cellIntStyle, cell, propValue);
            }
        }

//        for (int j = 0; j < excel.getColumnSize(); j++) {
//            sheet.autoSizeColumn(j);
//        }
    }

    //初始化excel文件设置
    public static ExcelFile initExcelFile(XlsWriterSetting setting){
        ExcelFile file = new ExcelFile();
        file.setSetting(setting);

        Workbook wb;
        if(setting.getType().equals(XlsType.XLS)) {
            wb = new HSSFWorkbook();
        }else if(setting.isUseStreamWriter()){
            wb = new SXSSFWorkbook(100);
        }else{
            wb = new XSSFWorkbook();
        }
        file.setWorkbook(wb);

        int rowIndex = setting.getStartRow();

        List<CellWriterSetting> columnList = setting.getColumnList();
        int columnSize = columnList.size();

        Sheet sheet = getSheet(file);
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

        CellStyle cellIntStyle = wb.createCellStyle();
        cellIntStyle.cloneStyleFrom(cellBorderStyle);
        cellIntStyle.setDataFormat(createHelper.createDataFormat().getFormat("0"));

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

        file.setCurrentRow(rowIndex);
        file.setColumnSize(columnSize);
        file.setCellBorderStyle(cellBorderStyle);
        file.setCellDateStyle(cellDateStyle);
        file.setCellDoubleStyle(cellDoubleStyle);
        file.setCellIntStyle(cellIntStyle);
        file.setCurrentIndexNum(setting.getIndexStartValue());

        return file;
    }

    private static Sheet getSheet(ExcelFile excel) {
        Workbook wb = excel.getWorkbook();
        XlsWriterSetting setting = excel.getSetting();
        String safeName = WorkbookUtil.createSafeSheetName(setting.getSheetName());
        Sheet sheet = wb.getSheet(safeName);
        if(sheet == null) {
            sheet = wb.createSheet(safeName);
        }
        return sheet;
    }

    private static void setCellValue(CellStyle cellDateStyle, CellStyle cellDoubleStyle, CellStyle cellIntStyle, Cell cell, Object propValue) {
        if(propValue instanceof Number){
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);

            if(propValue instanceof Integer){
                cell.setCellStyle(cellIntStyle);
                cell.setCellValue((Integer) propValue);
            } else if(propValue instanceof Double){
                cell.setCellStyle(cellDoubleStyle);
                cell.setCellValue((Double) propValue);
            } else if(propValue instanceof Float){
                cell.setCellStyle(cellDoubleStyle);
                cell.setCellValue((Float) propValue);
            } else if (propValue instanceof Long) {
                cell.setCellStyle(cellIntStyle);
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


