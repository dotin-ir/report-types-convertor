package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public class XLSColumnDefinitionBuilder {

    private int defaultColWidth = 2;
    private List<XLSColumnDefinition> columnDefinitions;
    private List<XLSColumnDefinition> tmpColumnDefinitions;
    private List<String> tmpUniqueColumns;
    private List<String> optionalColumns;

    public XLSColumnDefinitionBuilder() {
        columnDefinitions = new ArrayList<XLSColumnDefinition>();
        tmpColumnDefinitions = new ArrayList<XLSColumnDefinition>();
        tmpUniqueColumns = new ArrayList<String>();
        optionalColumns = new ArrayList<String>();
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName) {
        synchronized (columnDefinitions) {
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, defaultColWidth);
            columnDefinitions.add(columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, boolean hidden) {
        synchronized (columnDefinitions) {
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, defaultColWidth, hidden);
            columnDefinitions.add(columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width) {
        synchronized (columnDefinitions) {
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, width);
            columnDefinitions.add(columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width, boolean hidden) {
        synchronized (columnDefinitions) {
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName, width, hidden);
            columnDefinitions.add(columnDefinition);
            return this;
        }
    }

    public List<XLSColumnDefinition> build() {
        synchronized (columnDefinitions) {
            Map<String, String> cols = new HashMap<String, String>();
            for (XLSColumnDefinition definition : columnDefinitions) {

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
                    for (XLSColumnDefinition subDefinition : definition.getSubColumns()) {
                        if (cols.get(subDefinition.getName()) == null) {
                            cols.put(subDefinition.getName(), "");
                        } else {
                            throw new RuntimeException("Column with the same key [ " + subDefinition.getName() + " ] exist!");
                        }
                    }
                }
            }
        }
        return columnDefinitions;
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, List<XLSColumnDefinition> subCoumns) {
        synchronized (columnDefinitions) {
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(columnDefinitions.size(), name, fName);
            columnDefinition.setRealColumn(false);
            columnDefinition.setSubColumns(subCoumns);
            columnDefinitions.add(columnDefinition);
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
}
