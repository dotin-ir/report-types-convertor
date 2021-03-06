package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.XLSRowCellsData;
import org.apache.commons.lang.StringUtils;

import java.util.*;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public class XLSRowCellsDataBuilder {
    private Map<String, List<Map<String, String>>> row;
    private int maxRowLength;

    public XLSRowCellsDataBuilder() {
        initBuilder();
    }

    public XLSRowCellsDataBuilder addCellSimpleData(String key, String value) {
        key = checkKey(key);
        List<Map<String, String>> convertedValue = new ArrayList<Map<String, String>>();
        Map<String, String> map = new HashMap<String, String>();
        map.put(key, value);
        convertedValue.add(map);
        row.put(key, convertedValue);
        return this;
    }

    private String checkKey(String key) {
        if (StringUtils.isEmpty(key)){
            throw new RuntimeException("Key parameter is empty!");
        }
        key = key.trim();
        return key;
    }

    public XLSRowCellsDataBuilder addCellListData(String key, List<String> values) {
        key = checkKey(key);
        List<Map<String, String>> convertedValue = new ArrayList<Map<String, String>>();
        for (String value : values) {
            Map<String, String> map = new HashMap<String, String>();
            map.put(key, value);
            convertedValue.add(map);
            row.put(key, convertedValue);
        }
        if (maxRowLength < values.size()) {
            maxRowLength = values.size();
        }
        return this;
    }

    public XLSRowCellsDataBuilder addCellSimpleData(String key, List<String> values) {
        key = checkKey(key);
        List<Map<String, String>> convertedValue = new ArrayList<Map<String, String>>();
        for (String value : values) {
            Map<String, String> map = new HashMap<String, String>();
            map.put(key, value);
            convertedValue.add(map);
        }
        row.put(key, convertedValue);
        if (maxRowLength < values.size()) {
            maxRowLength = values.size();
        }
        return this;
    }

    public XLSRowCellsDataBuilder addEntityListCellsData(String key, List<Map<String, String>> values) {
        key = checkKey(key);
        row.put(key, values);
        if (maxRowLength < values.size()) {
            maxRowLength = values.size();
        }
        return this;
    }

    public XLSRowCellsDataBuilder addEntityCellsData(String key, Map<String, String> values) {
        key = checkKey(key);
        row.put(key, Arrays.asList(values));
        return this;
    }


    public XLSRowCellsData build() {
        XLSRowCellsData rowDefinition = new XLSRowCellsData(row, maxRowLength);
        initBuilder();
        return rowDefinition;
    }

    private void initBuilder() {
        maxRowLength = 1;
        row = new HashMap<String, List<Map<String, String>>>();
    }


}
