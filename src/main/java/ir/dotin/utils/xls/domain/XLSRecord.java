package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.checker.XLSUtils;
import org.apache.commons.lang.StringUtils;

import java.io.Serializable;
import java.util.*;

/**
 * Created by r.rastakfard on 6/22/2016.
 */
public class XLSRecord implements Comparable<XLSRecord>, Serializable {

    //        private boolean isValid;
    Map<String, List<Map<String, String>>> recordData;
    private List<XLSColumnDefinition> recordColumns;

    /*public XLSRecord(Map<String,List<Map<String, String>>> recordData) {
        this.recordData = recordData;
    }*/

    public XLSRecord() {
    }

    public XLSRecord(Map<String, String> record) {
        for (String key : record.keySet()) {
            String value = record.get(key);
            if (StringUtils.isEmpty(value)) {
                value = "";
            }
            Map<String, String> dummyRecord = new HashMap<String, String>();
            dummyRecord.put(key, value);
            List<Map<String, String>> valueStrings = Arrays.asList(dummyRecord);
            getRecordData().put(key, valueStrings);
        }

    }

    public Map<String, List<Map<String, String>>> getRecordData() {
        if (recordData == null) {
            recordData = new HashMap<String, List<Map<String, String>>>();
        }
        return recordData;
    }

    public void setRecordData(Map<String, List<Map<String, String>>> recordData) {
        this.recordData = recordData;
    }

    public int compareTo(XLSRecord o) {
        for (XLSColumnDefinition definition : getRecordColumns()) {
            if (!definition.isHidden()) {
                List<Map<String, String>> colListData = getRecordData().get(definition.getName());
                List<Map<String, String>> colListData1 = o.getRecordData().get(definition.getName());
                if (colListData == null) colListData = new ArrayList<Map<String, String>>();
                if (colListData1 == null) colListData1 = new ArrayList<Map<String, String>>();
                if (colListData.size() == colListData1.size()) {
                    for (int rowIndex = 0; rowIndex < colListData.size(); rowIndex++) {
                        int compareSimpleRecord = XLSUtils.compareSimpleRecord(colListData1.get(rowIndex), colListData1.get(rowIndex));
                        if (compareSimpleRecord != 0) {
                            return compareSimpleRecord;
                        }
                    }
                } else {
                    return new Integer(colListData.size()).compareTo(colListData1.size());
                }
            }
        }
        return 0;
    }


    public String getSimpleColumnValue(String colName) {
        Map<String, List<Map<String, String>>> recordData = getRecordData();
        List<Map<String, String>> colDataList = recordData.get(colName);
        if (colDataList != null && !colDataList.isEmpty()) {
            return colDataList.get(0).get(colName);
        }
        return null;
    }

    public String getSimpleColumnValue(String parentColName, String colName) {
        Map<String, List<Map<String, String>>> recordData = getRecordData();
        List<Map<String, String>> colDataList = recordData.get(parentColName);
        if (colDataList != null && !colDataList.isEmpty()) {
            return colDataList.get(0).get(colName);
        }
        return null;
    }

    public void setSimpleDataValue(String colName, String value) {
        Map<String, String> valueMap = new HashMap<String, String>();
        valueMap.put(colName, value);
        getRecordData().put(colName, Arrays.asList(valueMap));
    }

    public List<XLSColumnDefinition> getRecordColumns() {
        if (recordColumns == null) {
            recordColumns = new ArrayList<XLSColumnDefinition>();
        }
        return recordColumns;
    }

    public void setRecordColumns(List<XLSColumnDefinition> recordColumns) {
        this.recordColumns = recordColumns;
    }


}
