package ir.dotin.utils.xls.domain;

import org.apache.commons.lang.builder.ToStringBuilder;

import java.io.Serializable;
import java.util.List;

/**
 * Created by r.rastakfard on 7/10/2016.
 */
public class XLSReportField<E>  implements Serializable {
    private String name;
    private String value;
    private List<E> values;
    private int width = 1;

    public XLSReportField(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public XLSReportField(String name, List<E> values) {
        this.name = name;
        this.values = values;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public List<E> getValues() {
        return values;
    }

    public void setValues(List<E> values) {
        this.values = values;
    }

    @Override
    public String toString() {
        return new ToStringBuilder(this)
                .append("name", name)
                .append("value", value)
                .append("values", values)
                .append("width", width)
                .toString();
    }
}
