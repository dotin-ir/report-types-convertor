package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import org.apache.commons.lang.StringUtils;

import java.util.*;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public class XLSColumnDefinitionBuilder<B> {

    private int defaultColWidth = 2;
    private Map<String, XLSColumnDefinition> columnDefinitions;
    private List<XLSColumnDefinition> tmpColumnDefinitions;
    private List<String> tmpUniqueColumns;
    private List<String> optionalColumns;
    private int orderIndex = 0;

    public XLSColumnDefinitionBuilder() {
        columnDefinitions = new HashMap<String, XLSColumnDefinition>();
        tmpColumnDefinitions = new ArrayList<XLSColumnDefinition>();
        tmpUniqueColumns = new ArrayList<String>();
        optionalColumns = new ArrayList<String>();
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(orderIndex, name, fName, defaultColWidth);
            orderIndex++;
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
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(orderIndex, name, fName, defaultColWidth, hidden);
            orderIndex++;
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(orderIndex, name, fName, width);
            orderIndex++;
            columnDefinitions.put(name, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, Integer width, boolean hidden) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(orderIndex, name, fName, width, hidden);
            orderIndex++;
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
        List<XLSColumnDefinition> result = new ArrayList(columnDefinitions.values());
        Collections.sort(result, new Comparator<XLSColumnDefinition>() {
            public int compare(XLSColumnDefinition o1, XLSColumnDefinition o2) {
                return o1.getOrder().compareTo(o2.getOrder());
            }
        });
        return result;
    }

    public XLSColumnDefinitionBuilder addColumnDefinition(String name, String fName, List<XLSColumnDefinition> subCoumns) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(name);
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(defaultColWidth, name, fName);
            columnDefinition.setOrder(orderIndex);
            orderIndex++;
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
            columnDefinition.setOrder(orderIndex);
            orderIndex++;
            columnDefinition.setRealColumn(true);
            columnDefinition.setBasicInfoCollectionKey(basicInfoKey);
            columnDefinitions.put(colName, columnDefinition);
            return this;
        }
    }

    public XLSColumnDefinitionBuilder addColumnDefinitionWithBasicInfo(String colName, String fName, List<B> basicInfoList) {
        return addColumnDefinitionWithBasicInfo(colName, fName, colName, basicInfoList);
    }

    public XLSColumnDefinitionBuilder addColumnDefinitionWithBasicInfo(String colName, String fName, String basicInfoKey, List<B> basicInfoCollection) {
        synchronized (columnDefinitions) {
            checkColumnDuplication(colName);
            if (StringUtils.isEmpty(basicInfoKey)) {
                throw new IllegalArgumentException("Basic info key is empty!");
            }
            XLSColumnDefinition columnDefinition = new XLSColumnDefinition(defaultColWidth, colName, fName);
            columnDefinition.setOrder(orderIndex);
            orderIndex++;
            columnDefinition.setRealColumn(true);
            columnDefinition.setBasicInfoCollection(basicInfoCollection);
            columnDefinition.setBasicInfoCollectionKey(basicInfoKey);
            columnDefinitions.put(colName, columnDefinition);
            return this;
        }
    }
}
