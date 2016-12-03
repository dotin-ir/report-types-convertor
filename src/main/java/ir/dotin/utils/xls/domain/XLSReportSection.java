package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.renderer.XLSSectionRenderer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.Serializable;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public abstract class XLSReportSection implements Serializable {

    private List<XLSColumnDefinition> headerCols;
    private String emptyRecordsMessage = XLSConstants.DEFAULT_EMPTY_SECTION_RECORDS_MESSAGE;
    private Map<String, XLSBasicInfo> basicInfos;
    private HSSFWorkbook workBook;

    public List<XLSColumnDefinition> getHeaderCols() {
        return headerCols;
    }

    public void setHeaderCols(List<XLSColumnDefinition> headerCols) {
        this.headerCols = headerCols;
    }

    public abstract XLSSectionRenderer getRenderer();

    public abstract Integer getRecordsCount();

    public String getEmptyRecordsMessage() {
        return emptyRecordsMessage;
    }

    public void setEmptyRecordsMessage(String emptyRecordsMessage) {
        this.emptyRecordsMessage = emptyRecordsMessage;
    }

    public void setBasicInfos(Map<String, XLSBasicInfo> basicInfos) {
        this.basicInfos = basicInfos;
    }

    public Map<String, XLSBasicInfo> getBasicInfos() {
        if (basicInfos==null){
            basicInfos = new HashMap<String, XLSBasicInfo>();
        }
        return basicInfos;
    }

    public HSSFWorkbook getWorkBook() {
        return workBook;
    }

    public void setWorkBook(HSSFWorkbook workBook) {
        this.workBook = workBook;
    }
}
