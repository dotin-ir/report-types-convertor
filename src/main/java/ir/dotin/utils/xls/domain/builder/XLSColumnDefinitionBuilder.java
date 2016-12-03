package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import org.apache.commons.lang.StringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public class XLSColumnDefinitionBuilder<B> {

    private int defaultColWidth = 2;
    private Map<String, XLSColumnDefinition> columnDefinitions;
    private List<XLSColumnDefinition> tmpColumnDefinitions;
    private List<String> tmpUniqueColumns;
    private List<String> optionalColumns;

    public XLSColumnDefinitionBuilder() {
        columnDefinitions = new HashMap<String, XLSColumnDefinition>();
        tmpColumnDefinitions = new ArrayList<XLSColumnDefinition>();
        tmpUniqueColumns = new ArrayList<String>();
        optionalColumns = new ArrayList<String>();
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, defaultColWidth);
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    private void checkColumnDuplication(String name) {
        if (StringUtils.isNotEmpty(name) && columnDefinitions.containsKey(name)) {
            throw new RuntimeException("Duplicate column with name '" + name + "' found!");
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, boolean hidden) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, defaultColWidth, hidden);
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, width);
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width, boolean hidden) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, width, hidden);
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public List<XLSColumnDefinition> build() {
        synchronized (columnDefinitions) {
            Map<String, String> cols = new HashMap<String, String>();
            for (String colName : columnDefinitions.keySet()) {
                XLSColumnDefinition definition = columnDefinitions.get(colName);
                if (tmpUniqueColumns.contains(definition.getName())) {
                    if (definition.isRealColumn()) {
                        definition.setUniqueColumn(true);
                    } else {
                        throw new RuntimeException("you can not set a header column (with sub columns) as unique key!");
                    }
                }

                if (optionalColumns.contains(definition.getName())) {
                    definition.setMandatory(false);
                }

                if (cols.get(definition.getName()) == null) {
                    cols.put(definition.getName(), "");
                } else {
                    throw new RuntimeException("Column with the same key [ " + definition.getName() + " ] exist!");
                }
                if (!definition.getSubColumns().isEmpty()) {
                    List<XLSColumnDefinition> subColumns = definition.getSubColumns();
                    for (XLSColumnDefinition subDefinition : subColumns) {
                        if (cols.get(subDefinition.getName()) == null) {
                            cols.put(subDefinition.getName(), "");
                        } else {
                            throw new RuntimeException("Column with the same key [ " + subDefinition.getName() + " ] exist!");
                        }
                    }
                }
            }
        }
        return new ArrayList(columnDefinitions.values());
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, List<XLSColumnDefinition> subCoumns) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(defaultColWidth, name, fName);
            columnDefinition.setRealColumn(false);
            columnDefinition.setSubColumns(subCoumns);
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder setDefaultColumnWidth(int width) {
        if (width > 0) {
            defaultColWidth = width;
        }
        return this;
    }


    public XLSColumnDefinitionBuilder addUniqueColumn(String colName) {
        if (!tmpUniqueColumns.contains(colName)) {
            tmpUniqueColumns.add(colName);
        }
        return this;
    }

    public XLSColumnDefinitionBuilder addOptionalColumn(String colName) {
        if (!optionalColumns.contains(colName)) {
            optionalColumns.add(colName);
        }
        return this;
    }

    public XLSColumnDefinitionBuilder addColumnDefinitionWithBasicInfo(String colName, String fName, String basicInfoKey) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(colName);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(defaultColWidth, colName, fName);
            columnDefinition.setRealColumn(true);
            columnDefinition.setBasicInfoCollectionKey(basicInfoKey);
            columnDefinitions.put(colName, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinitionWithBasicInfo(String colName, String fName, List<B> basicInfoList) {
        return addColumnDefinitionWithBasicInfo(colName,fName,colName,basicInfoList);
    }

    public XLSColumnDefinitionBuilder addColumnDefinitionWithBasicInfo(String colName, String fName, String basicInfoKey, List<B> basicInfoCollection) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(colName);
            if (StringUtils.isEmpty(basicInfoKey)) {
                throw new IllegalArgumentException("Basic info key is empty!");
            }
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(defaultColWidth, colName, fName);
            columnDefinition.setRealColumn(true);
            columnDefinition.setBasicInfoCollection(basicInfoCollection);
            columnDefinition.setBasicInfoCollectionKey(basicInfoKey);
            columnDefinitions.put(colName, columnDefinition);
            return this;
        }
    }
}
