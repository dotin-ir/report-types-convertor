package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.mapper.XLSEntityToRowMapper;
import ir.dotin.utils.xls.renderer.XLSListSectionRenderer;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import ir.dotin.utils.xls.renderer.XLSSectionRenderer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/10/2016.
 */
public class XLSListReportSection<E> extends XLSReportSection {

    private String title;
    private List<XLSReportField> titleFields;
    private List<Map<String, String>> rawRecords;
    private List<E> records;
    private XLSRowCustomizer rowCustomizer;
    private XLSEntityToRowMapper entityToRowMapper;
    private XLSListSectionRenderer sectionRenderer = new XLSListSectionRenderer();

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<XLSReportField> getTitleFields() {
        if (titleFields == null) {
            titleFields = new ArrayList<XLSReportField>();
        }
        return titleFields;
    }

    public void setTitleFields(List<XLSReportField> titleFields) {
        this.titleFields = titleFields;
    }

    public List<Map<String, String>> getRawRecords() {
        if (rawRecords == null) {
            rawRecords = new ArrayList<Map<String, String>>();
        }
        return rawRecords;
    }

    public void setRawRecords(List<Map<String, String>> rawRecords) {
        this.rawRecords = rawRecords;
    }

    public XLSRowCustomizer getRowCustomizer() {
        return rowCustomizer;
    }

    public void setRowCustomizer(XLSRowCustomizer rowCustomizer) {
        this.rowCustomizer = rowCustomizer;
    }

    public List<E> getRecords() {
        if (records == null) {
            records = new ArrayList<E>();
        }
        return records;
    }

    public void setRecords(List<E> records) {
        this.records = records;
    }

    public void addRecord(E record) {
        if (record != null) {
            getRecords().add(record);
        }
    }

    public XLSEntityToRowMapper getEntityToRowMapper() {
        return entityToRowMapper;
    }

    public void setEntityToRowMapper(XLSEntityToRowMapper entityToRowMapper) {
        this.entityToRowMapper = entityToRowMapper;
    }

    public XLSSectionRenderer getRenderer() {
        return sectionRenderer;
    }

    public Integer getRecordsCount() {
        List<E> records = getRecords();
        List<Map<String, String>> rawRecords = getRawRecords();
        return records.isEmpty() ? rawRecords.size() : records.size();
    }

    public void addTitleField(XLSReportField reportField) {
        if (reportField == null) {
            throw new IllegalArgumentException("report field can not be null!");
        }
        getTitleFields().add(reportField);
    }


}
