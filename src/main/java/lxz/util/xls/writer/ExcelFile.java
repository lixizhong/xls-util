package lxz.util.xls.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by lixizhong on 2017/1/4.
 */
public class ExcelFile {
    private Workbook workbook;
    private CellStyle cellBorderStyle;
    private CellStyle cellDateStyle;
    private CellStyle cellDoubleStyle;
    private CellStyle cellIntStyle;
    private XlsWriterSetting setting;
    private int currentRow; //当前行
    private int currentIndexNum; //当前序号
    private int columnSize;

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public CellStyle getCellBorderStyle() {
        return cellBorderStyle;
    }

    public void setCellBorderStyle(CellStyle cellBorderStyle) {
        this.cellBorderStyle = cellBorderStyle;
    }

    public CellStyle getCellDateStyle() {
        return cellDateStyle;
    }

    public void setCellDateStyle(CellStyle cellDateStyle) {
        this.cellDateStyle = cellDateStyle;
    }

    public CellStyle getCellDoubleStyle() {
        return cellDoubleStyle;
    }

    public void setCellDoubleStyle(CellStyle cellDoubleStyle) {
        this.cellDoubleStyle = cellDoubleStyle;
    }

    public CellStyle getCellIntStyle() {
        return cellIntStyle;
    }

    public void setCellIntStyle(CellStyle cellIntStyle) {
        this.cellIntStyle = cellIntStyle;
    }

    public XlsWriterSetting getSetting() {
        return setting;
    }

    public void setSetting(XlsWriterSetting setting) {
        this.setting = setting;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public int getAndPlusCurrentRow() {
        return currentRow++;
    }

    public void setCurrentRow(int currentRow) {
        this.currentRow = currentRow;
    }

    public int getColumnSize() {
        return columnSize;
    }

    public void setColumnSize(int columnSize) {
        this.columnSize = columnSize;
    }

    public int getCurrentIndexNum() {
        return currentIndexNum;
    }

    public int getCurrentAndPlusIndexNum() {
        return currentIndexNum++;
    }

    public void setCurrentIndexNum(int currentIndexNum) {
        this.currentIndexNum = currentIndexNum;
    }
}
