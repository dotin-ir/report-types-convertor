package ir.dotin.utils.xls.domain;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;

import java.io.Serializable;

/**
 * Created by r.rastakfard on 7/17/2016.
 */
public class XLSColorDescription implements Serializable {
    private String key;
    private int red;
    private int green;
    private int blue;
    private String description;
    private short colorIndex;
    private boolean defaultColor = false;
    private int realIndex;

    private XLSColorDescription() {
    }

    public XLSColorDescription(String key, int red, int green, int blue, String description) {
        this.key = key;
        this.red = red;
        this.green = green;
        this.blue = blue;
        this.description = description;
        validate();
    }

    public int getRed() {
        return red;
    }

    public void setRed(int red) {
        this.red = red;
        validate();
    }

    public int getGreen() {
        return green;
    }

    public void setGreen(int green) {
        this.green = green;
        validate();
    }

    public int getBlue() {
        return blue;
    }

    public void setBlue(int blue) {
        this.blue = blue;
        validate();
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public void validate() {
        if (StringUtils.isEmpty(key)){
            throw new IllegalArgumentException("Color key is empty!");
        }
        if (red < 0 || red > 255) {
            throw new IllegalArgumentException("Red must be between 0..255");
        }
        if (green < 0 || green > 255) {
            throw new IllegalArgumentException("Green must be between 0..255");
        }
        if (blue < 0 || blue > 255) {
            throw new IllegalArgumentException("Blue must be between 0..255");
        }
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public short getColorIndex() {
        return colorIndex;
    }

    public void setColorIndex(short colorIndex) {
        this.colorIndex = colorIndex;
    }

    public boolean isDefaultColor() {
        return defaultColor;
    }

    public void setDefaultColor(boolean defaultColor) {
        this.defaultColor = defaultColor;
    }

    public void setRealIndex(int realIndex) {
        this.realIndex = realIndex;
    }

    public int getRealIndex() {
        return realIndex;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;

        if (!(o instanceof XLSColorDescription)) return false;

        XLSColorDescription that = (XLSColorDescription) o;

        return new EqualsBuilder()
                .append(key, that.key)
                .isEquals();
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder(17, 37)
                .append(key)
                .toHashCode();
    }
}
