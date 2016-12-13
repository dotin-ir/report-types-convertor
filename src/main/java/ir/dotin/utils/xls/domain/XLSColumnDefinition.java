package ir.dotin.utils.xls.domain;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public class XLSColumnDefinition<B> implements Serializable {
    public static short SUM_AGGREGATION = 0;
    public static short AVG_AGGREGATION = 1;

    private boolean hidden = false;
    private Integer order = Integer.MAX_VALUE;
    private String name;
    private String fName;
    /* based on character*/
    private Integer width;
    private String defaultEmptyData = "";
    private short color = IndexedColors.AQUA.getIndex();
    private List<XLSColumnDefinition> subColumns;
    private short aggregationType = -1;
    private Integer totalSubColumnsWidth = 0;
    private boolean isRealColumn = true;
    private boolean isUniqueColumn = false;
    private boolean mandatory = true;
    private Class<?> fieldType = String.class;
    private String basicInfoCollectionKey;
    private List<B> basicInfoCollection;

    public XLSColumnDefinition() {
    }

    public XLSColumnDefinition(int order, String name, String fName, short color) {
        if (StringUtils.isEmpty(name)){
            throw new IllegalArgumentException("Column name is Empty!");
        }
        if (StringUtils.isEmpty(fName)){
            throw new IllegalArgumentException("Column fName is Empty!");
        }
        this.order = order;
        this.name = name;
        this.fName = fName;
        this.color = color;
    }

    public XLSColumnDefinition(int order, String name, String fName, Integer width) {
        this(width,name,fName);
        this.order = order;
    }

    public XLSColumnDefinition(int order, String name, String fName, Integer width, boolean hidden) {
        this(order, name, fName, width);
        this.hidden = hidden;
    }

    public XLSColumnDefinition(int width, String name, String fName) {
        if (StringUtils.isEmpty(name)){
            throw new IllegalArgumentException("Column name is Empty!");
        }
        if (StringUtils.isEmpty(fName)){
            throw new IllegalArgumentException("Column fName is Empty!");
        }
        this.name = name.trim();
        this.width = width;
        this.fName = fName.trim();
    }

    public Integer getOrder() {
        return order;
    }

    public void setOrder(Integer order) {
        this.order = order;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getfName() {
        return fName;
    }

    public void setfName(String fName) {
        this.fName = fName;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        if (width < 1) {
            throw new RuntimeException("column width must greater than 0");
        }
        this.width = width;
    }

    public boolean isHidden() {
        return hidden;
    }

    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    public short getAggregationType() {
        return aggregationType;
    }

    public void setAggregationType(short aggregationType) {
        this.aggregationType = aggregationType;
    }

    public List<XLSColumnDefinition> getSubColumns() {
        if (subColumns == null) {
            subColumns = new ArrayList<XLSColumnDefinition>();
        }
        return subColumns;
    }

    public void setSubColumns(List<XLSColumnDefinition> subColumns) {
        if (subColumns == null || subColumns.isEmpty()) return;
        setRealColumn(false);
        this.subColumns = subColumns;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public Integer getTotalSubColumnsWidth() {
        if (totalSubColumnsWidth == 0) {
            for (XLSColumnDefinition definition : getSubColumns()) {
                totalSubColumnsWidth += definition.getWidth();
            }
        }
        return totalSubColumnsWidth;
    }

    public Integer getRealColumnWidth() {
        Integer width;
        if (!getSubColumns().isEmpty()) {
            width = getTotalSubColumnsWidth();
        } else {
            width = getWidth();
        }
        return width;
    }

    public String getDefaultEmptyData() {
        return defaultEmptyData;
    }

    public void setDefaultEmptyData(String defaultEmptyData) {
        this.defaultEmptyData = defaultEmptyData;
    }


    public boolean isRealColumn() {
        return isRealColumn;
    }

    public void setRealColumn(boolean realColumn) {
        isRealColumn = realColumn;
    }

    public boolean isUniqueColumn() {
        return isUniqueColumn;
    }

    public void setUniqueColumn(boolean uniqueColumn) {
        isUniqueColumn = uniqueColumn;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;

        if (o == null || getClass() != o.getClass()) return false;

        XLSColumnDefinition that = (XLSColumnDefinition) o;

        return new EqualsBuilder()
                .append(name, that.name)
                .append(fName, that.fName)
                .isEquals();
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder(17, 37)
                .append(name)
                .append(fName)
                .toHashCode();
    }

    public boolean isMandatory() {
        return mandatory;
    }

    public void setMandatory(boolean mandatory) {
        this.mandatory = mandatory;
    }

    @Override
    public String toString() {
        return new ToStringBuilder(this)
                .append("name", name)
                .append("fName", fName)
                .append("width", width)
                .toString();
    }

    public Class<?> getFieldType() {
        return fieldType;
    }

    public void setFieldType(Class<?> fieldType) {
        this.fieldType = fieldType;
    }

    public void setBasicInfoCollectionKey(String bascInfoCollectionKey) {
        this.basicInfoCollectionKey = bascInfoCollectionKey;
    }

    public String getBasicInfoCollectionKey() {
        return basicInfoCollectionKey;
    }

    public void setBasicInfoCollection(List<B> bascInfoCollection) {
        this.basicInfoCollection = bascInfoCollection;
    }

    public List<B> getBasicInfoCollection() {
        return basicInfoCollection;
    }

}
